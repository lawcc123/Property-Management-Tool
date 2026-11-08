"""Verified browser-free API sessions, with interactive authentication only when needed."""
from contextlib import contextmanager
import json
import os
from pathlib import Path
import tempfile
import time
from types import SimpleNamespace
import login_flow
from request_policy import Cooldown, guard_response


class SessionUnavailable(RuntimeError):
    pass


class CookieSource:
    def __init__(self, state):
        self.state = state

    def cookies(self):
        return self.state.get("cookies", [])


def load_state(path):
    try:
        state = json.loads(Path(path).read_text(encoding="utf-8-sig"))
        return state if isinstance(state, dict) and isinstance(state.get("cookies"), list) else None
    except (OSError, ValueError):
        return None


def make_client(playwright, state, user_agent=None):
    options = {"storage_state": {"cookies": state.get("cookies", []), "origins": state.get("origins", [])},
               "timeout": 30000, "max_redirects": 0, "ignore_https_errors": True}
    if isinstance(user_agent, str) and user_agent:
        options["user_agent"] = user_agent
    return playwright.request.new_context(**options)


def verify(client, sync, headers):
    """A single-row read-only IRGrid search; no auth redirects or raw body logging."""
    try:
        response = client.post(sync.IRGRID_URL,
            params={"NumRecords": 1, "search": "", "enddate": "", "invoicedatefrom": "",
                    "invoicedateto": "", "startdate": ""},
            data="[]", headers=headers, timeout=15000, max_redirects=0)
    except Exception:
        raise SessionUnavailable("Yardi could not be reached. Check your connection and try again; no workbook changes were saved.") from None
    try:
        guard_response(response, lambda: Cooldown.for_sync(sync), auth_probe=True)
        if response.status in (401, 403):
            return False
        if response.status in (301, 302, 303, 307, 308):
            location = response.headers.get("location", "").lower()
            if any(word in location for word in ("login", "signin", "sign-in", "account", "authorize")):
                return False
            raise SessionUnavailable("Yardi returned an unexpected redirect. No workbook changes were saved.")
        if not response.ok:
            raise SessionUnavailable(f"Yardi returned HTTP {response.status}. Try again later; no workbook changes were saved.")
        if "text/html" in response.headers.get("content-type", "").lower():
            return False
        try:
            body = response.json()
        except Exception:
            raise SessionUnavailable("Yardi returned an unreadable response. No workbook changes were saved.") from None
        if isinstance(body, list):
            return True
        if (isinstance(body, dict) and isinstance(body.get("Records"), list)
                and not any(body.get(key) for key in ("error", "Error", "Exception"))):
            return True
        raise SessionUnavailable("Yardi returned an unexpected invoice-list response. No workbook changes were saved.")
    finally:
        response.dispose()


def save_state(path, state, user_agent, emit):
    """Retain refreshed cookies and the authenticated browser's own user agent."""
    destination = Path(path)
    staged = None
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        state = {"cookies": state.get("cookies", []), "origins": state.get("origins", []),
                 "p2p_user_agent": user_agent if isinstance(user_agent, str) else ""}
        descriptor, staged = tempfile.mkstemp(prefix="auth-", suffix=".tmp", dir=destination.parent)
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(state, stream)
        os.replace(staged, destination)
    except OSError:
        emit("log", message="The verified session could not be remembered. This sync can continue; sign-in may be needed next time.")
    finally:
        if staged and os.path.exists(staged):
            os.unlink(staged)


@contextmanager
def open_session(playwright, sync, credentials, retry_file, emit, check_cancel):
    client = None
    browser = None
    user_agent = None
    try:
        check_cancel()
        Cooldown.for_sync(sync).check()
        emit("stage", stage="session_check", message="Checking your saved Yardi session in the background...")
        cached = load_state(sync.AUTH_STATE_PATH)
        if cached:
            headers = sync.get_auth_headers(CookieSource(cached))
            if headers:
                user_agent = cached.get("p2p_user_agent")
                try:
                    client = make_client(playwright, cached, user_agent)
                except Exception:
                    raise SessionUnavailable("The saved session could not be loaded. Clear the saved session or try again.") from None
                if verify(client, sync, headers):
                    check_cancel()
                    emit("login_state", state="connected", message="Saved session confirmed. Syncing without opening a browser.")
                    save_state(sync.AUTH_STATE_PATH, client.storage_state(), user_agent, emit)
                    yield SimpleNamespace(request=client), headers
                    save_state(sync.AUTH_STATE_PATH, client.storage_state(), user_agent, emit)
                    return
                client.dispose()
                client = None
        check_cancel()
        emit("stage", stage="sign_in", message="Yardi requires authentication. Opening the sign-in browser...")
        os.makedirs(sync.SESSION_DIR, exist_ok=True)
        try:
            browser = playwright.chromium.launch_persistent_context(sync.SESSION_DIR, channel="chrome", headless=False,
                                                                    ignore_https_errors=True)
        except Exception:
            browser = playwright.chromium.launch_persistent_context(sync.SESSION_DIR, headless=False,
                                                                    ignore_https_errors=True)
        sync.restore_auth_state(browser)
        page = browser.new_page()
        page.set_default_timeout(30000)
        next_probe = 0
        last_notice = None

        def confirm_background(headers):
            nonlocal client, user_agent, next_probe, last_notice
            check_cancel()
            if time.monotonic() < next_probe:
                return False
            next_probe = time.monotonic() + 2
            candidate = None
            try:
                try:
                    state = browser.storage_state()
                except Exception:
                    raise SessionUnavailable("Browser sign-in state is not ready yet.") from None
                try:
                    user_agent = next(p for p in browser.pages if not p.is_closed()).evaluate("navigator.userAgent")
                except Exception:
                    user_agent = None
                try:
                    candidate = make_client(playwright, state, user_agent)
                except Exception:
                    raise SessionUnavailable("The background session is not ready yet.") from None
                accepted = verify(candidate, sync, headers)
                check_cancel()
                if accepted:
                    try:
                        confirmed_state = candidate.storage_state()
                    except Exception:
                        raise SessionUnavailable("The background session could not be captured yet.") from None
                    save_state(sync.AUTH_STATE_PATH, confirmed_state, user_agent, emit)
                    client = candidate
                    candidate = None
                    return True
                message = "P2P background access is not confirmed yet. Complete verification or select the required database in the browser. It will stay open."
            except SessionUnavailable:
                message = "Yardi background access could not be confirmed yet. The browser will stay open; you can wait, retry sign-in or cancel."
            finally:
                if candidate is not None:
                    candidate.dispose()
            if message != last_notice:
                emit("login_state", state="handoff_wait", message=message)
                last_notice = message
            return False

        headers = login_flow.authenticate(sync, browser, page, credentials, retry_file, emit, check_cancel,
                                          confirm_session=confirm_background)
        if client is None:
            raise SessionUnavailable("Background access was not confirmed. No workbook changes were saved.")
        credentials = None
        check_cancel()
        # Authentication and a separate API client have both been verified.
        browser.close()
        browser = None
        emit("login_state", state="connected", message="P2P session confirmed. The browser is closed; sync continues in the background.")
        yield SimpleNamespace(request=client), headers
        save_state(sync.AUTH_STATE_PATH, client.storage_state(), user_agent, emit)
    finally:
        if browser is not None:
            browser.close()
        if client is not None:
            client.dispose()
