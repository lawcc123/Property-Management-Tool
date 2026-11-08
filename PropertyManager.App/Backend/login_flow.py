"""Normal login/Elevate actions; security verification always remains interactive."""
from pathlib import Path
import time
from urllib.parse import urlsplit


def origin(url):
    try:
        parsed = urlsplit(url)
        if parsed.scheme not in ("http", "https") or not parsed.hostname:
            return None
        return parsed.scheme.lower(), parsed.hostname.lower(), parsed.port or (443 if parsed.scheme == "https" else 80)
    except (ValueError, TypeError):
        return None


def try_autofill(page, login_url, credentials):
    """Fill only an unambiguous, same-origin, POST login form. Never submit it.

    Exceptions intentionally never escape: Playwright call logs can include values.
    `waiting` makes no edits; `manual` requires manual entry or a deliberate retry.
    """
    expected = origin(login_url)
    if expected is None or expected[0] != "https" or not credentials or origin(credentials.get("origin", "")) != expected:
        return "manual"
    try:
        login_path = urlsplit(login_url).path.rstrip("/").lower()
        if origin(page.url) != expected or urlsplit(page.url).path.rstrip("/").lower() != login_path:
            return "waiting"
        passwords = page.locator('input[type="password"]')
        visible = [passwords.nth(i) for i in range(passwords.count())
                   if passwords.nth(i).is_visible() and passwords.nth(i).is_enabled()]
        if len(visible) != 1:
            return "waiting"
        password = visible[0]
        if password.get_attribute("autocomplete") in ("new-password", "one-time-code"):
            return "manual"
        form = password.evaluate("el => el.form ? {action: el.form.action, method: el.form.method} : null")
        if (not form or form["method"].lower() != "post" or origin(form["action"]) != expected
                or urlsplit(form["action"]).path.rstrip("/").lower() != login_path):
            return "manual"
        # Select the username only within the same form as the password.
        username_handle = password.evaluate_handle('''el => {
            const candidates = Array.from(el.form.elements).filter(x =>
                x.tagName === 'INPUT' && !x.disabled && !x.readOnly &&
                ['text', 'email'].includes(x.type) && x.getClientRects().length > 0);
            const named = candidates.filter(x => x.autocomplete === 'username' || x.type === 'email' ||
                /^(username|user|email|login|emailaddress)$/i.test(x.name || '') ||
                /^(username|user|email|login|emailaddress)$/i.test(x.id || ''));
            return named.length === 1 ? named[0] : named.length === 0 && candidates.length === 1 ? candidates[0] : null;
        }''')
        try:
            username = username_handle.as_element()
            if username is None:
                return "waiting"
            # Do not overwrite values a person or a password manager already entered.
            entered_username = username.input_value()
            if (entered_username and entered_username != credentials["username"]) or password.input_value():
                return "manual"
            if not entered_username:
                username.fill(credentials["username"], timeout=2500)
            password.fill(credentials["password"], timeout=2500)
            return "filled"
        finally:
            username_handle.dispose()
    except Exception:
        return "manual"


def try_submit_login(page, login_url):
    """Click a recognized login submit control once; never Verify or an MFA form."""
    try:
        expected = origin(login_url)
        login_path = urlsplit(login_url).path.rstrip("/").lower()
        if (expected is None or expected[0] != "https" or origin(page.url) != expected
                or urlsplit(page.url).path.rstrip("/").lower() != login_path):
            return "waiting"
        fields = page.locator('input[type="password"]')
        fields = [fields.nth(i) for i in range(fields.count()) if fields.nth(i).is_visible()]
        if len(fields) != 1:
            return "waiting"
        password = fields[0]
        if password.get_attribute("autocomplete") in ("new-password", "one-time-code"):
            return "waiting"
        ready = password.evaluate("""el => {
            if (!el.form || !el.value || el.disabled || el.form.method.toLowerCase() !== 'post') return null;
            const controls = Array.from(el.form.elements);
            const usernames = controls.filter(x => x.tagName === 'INPUT' && ['text','email'].includes(x.type) &&
                x.getClientRects().length > 0 && !x.disabled && x.value);
            const verification = controls.filter(x => /^(cf-turnstile-response|g-recaptcha-response|h-captcha-response)$/.test(x.name));
            if (usernames.length !== 1 || verification.some(x => !x.value)) return null;
            return {action: el.form.action};
        }""")
        if not ready:
            return "waiting"
        if origin(ready["action"]) != expected or urlsplit(ready["action"]).path.rstrip("/").lower() != login_path:
            return "manual"
        handle = password.evaluate_handle("""el => {
            const buttons = Array.from(el.form.elements).filter(x =>
                (x.tagName === 'BUTTON' && x.type === 'submit' || x.tagName === 'INPUT' && x.type === 'submit') &&
                !x.disabled && x.getClientRects().length > 0 &&
                /^(log\\s*in|sign\\s*in)$/i.test((x.innerText || x.value || x.getAttribute('aria-label') || '').trim()));
            return buttons.length === 1 ? buttons[0] : null;
        }""")
        try:
            button = handle.as_element()
            if button is None:
                return "manual"
            # A timeout after this click may mean submission occurred. Do not retry it automatically.
            try:
                button.click(timeout=2500)
                return "submitted"
            except Exception:
                return "uncertain"
        finally:
            handle.dispose()
    except Exception:
        return "waiting"


def try_elevate(page, sync):
    """Select only the exact Production application on a configured origin."""
    import re
    try:
        allowed = {origin(sync.LOGIN_URL), origin(sync.P2P_ORIGIN)}
        if origin(page.url) not in allowed or origin(page.url) is None:
            return "waiting"
        label = re.compile(r"^Elevate\s*\(Production\)$", re.I)
        for role in ("button", "link", None):
            locator = page.get_by_role(role, name=label) if role else page.get_by_text(label, exact=True)
            matches = [locator.nth(i) for i in range(locator.count())
                       if locator.nth(i).is_visible() and locator.nth(i).is_enabled()]
            if len(matches) > 1:
                return "manual"
            if len(matches) == 1:
                try:
                    matches[0].click(timeout=2500)
                    return "clicked"
                except Exception:
                    return "uncertain"
        return "waiting"
    except Exception:
        return "waiting"


def authenticate(sync, context, page, credentials, retry_file, emit, check_cancel,
                 reminder_seconds=300, poll_seconds=0.3, confirm_session=None):
    """Wait until authentication succeeds, the user cancels, or closes the browser.

    Navigation failures and the five-minute reminder never close the browser.
    Retry is an explicit user action delivered via a unique per-run control file.
    """
    headers = sync.get_auth_headers(context)
    if headers and (confirm_session is None or confirm_session(headers)):
        emit("login_state", state="connected", message="Using your confirmed Yardi session.")
        return headers

    last_retry = None
    attempted = False
    submitted = False
    filled_here = False
    elevated = False
    notice = None
    reminder_at = time.monotonic() + reminder_seconds

    def notify(state, message):
        nonlocal notice
        if notice != (state, message):
            emit("login_state", state=state, message=message)
            notice = state, message

    def navigate():
        try:
            page.goto(sync.LOGIN_URL, wait_until="domcontentloaded", timeout=15000)
            notify("waiting", "Opening the normal sign-in page. Login and Elevate will be selected when ready; complete verification or MFA yourself.")
        except Exception:
            notify("retry_available", "The sign-in page did not finish loading. The browser is still open; finish there or choose Retry sign-in.")

    navigate()
    while True:
        check_cancel()
        headers = sync.get_auth_headers(context)
        if headers and (confirm_session is None or confirm_session(headers)):
            notify("connected", "P2P background access confirmed. Continuing sync.")
            return headers
        try:
            pages = [candidate for candidate in context.pages if not candidate.is_closed()]
        except Exception:
            pages = []
        if not pages:
            raise RuntimeError("The sign-in browser was closed. Start sync again when you are ready to sign in.")
        if page.is_closed():
            page = pages[-1]
        if retry_file:
            try:
                token = Path(retry_file).read_text(encoding="utf-8-sig").strip()
            except OSError:
                token = None
            if token and token != last_retry:
                last_retry = token
                attempted = submitted = elevated = filled_here = False
                reminder_at = time.monotonic() + reminder_seconds
                # Reload only when the user explicitly requests it.
                navigate()
        if credentials and not attempted:
            for candidate in pages:
                result = try_autofill(candidate, sync.LOGIN_URL, credentials)
                if result == "filled":
                    attempted = filled_here = True
                    notify("filled", "Credentials filled. Submitting login once when the normal Sign in button is ready.")
                    break
                if result == "manual":
                    attempted = True
                    notify("manual", "Please finish sign-in in the browser. Autofill did not overwrite existing fields or use an unrecognized form.")
                    break
        if filled_here and not submitted:
            for candidate in pages:
                result = try_submit_login(candidate, sync.LOGIN_URL)
                if result in ("submitted", "uncertain", "manual"):
                    submitted = True
                    if result == "submitted":
                        notify("submitted", "Login submitted once. Enter any one-time passcode and click Verify yourself; the browser will stay open.")
                    else:
                        notify("manual", "Finish sign-in in the browser. The app will not submit the login repeatedly.")
                    break
        if not elevated:
            for candidate in pages:
                result = try_elevate(candidate, sync)
                if result in ("clicked", "uncertain", "manual"):
                    elevated = True
                    if result == "clicked":
                        notify("elevate", "Opening Elevate (Production). The browser will stay open until P2P background access is confirmed.")
                    else:
                        notify("manual", "Please select Elevate (Production) in the browser if it has not opened. The app will not click it repeatedly.")
                    break
        if time.monotonic() >= reminder_at:
            notify("retry_available", "Still waiting for sign-in. The browser will stay open until you finish or cancel; Retry sign-in reloads the login page.")
            reminder_at = time.monotonic() + reminder_seconds
        # Pump Playwright's events while waiting; never click verification controls.
        try:
            page.wait_for_timeout(poll_seconds * 1000)
        except Exception:
            check_cancel()
            time.sleep(min(poll_seconds, 0.3))
