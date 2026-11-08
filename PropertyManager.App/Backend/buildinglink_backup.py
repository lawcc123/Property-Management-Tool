"""Unattended BuildingLink Occupant Data and Unit Data backup.

Each corporation is processed in its own clean Microsoft Edge session. Credentials
are read from Windows Credential Manager and are never written to configuration or
log files.
"""

from __future__ import annotations

import argparse
import getpass
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from contextlib import contextmanager
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import keyring
from keyring.errors import KeyringError, PasswordDeleteError
from playwright.sync_api import BrowserContext, Error as PlaywrightError, Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


APP_NAME = "BuildingLinkDataBackup"
CREDENTIAL_SERVICE = "BuildingLink.DataBackup"
EXIT_SUCCESS = 0
EXIT_AUTOMATION = 10
EXIT_CREDENTIALS = 11
EXIT_CONFIG = 15
DEFAULT_ACTION_TIMEOUT_SECONDS = 20
DEFAULT_LOGIN_TIMEOUT_SECONDS = 60
DEFAULT_DOWNLOAD_TIMEOUT_SECONDS = 60
DEFAULT_MAX_ATTEMPTS = 3
UI_SETTLE_MS = 750

PORTAL_URL = "https://secure.buildinglink.com/"
OCCUPANT_REPORT_URL = (
    "https://secure.buildinglink.com/reports/staff/download-data/"
    "download-occupantData-requests"
)
UNIT_REPORT_URL = (
    "https://secure.buildinglink.com/reports/staff/download-data/"
    "download-unit-profile-data"
)
LEGACY_UNIT_REPORT_PATH = "/v2/mgmt/export/unitsearch.aspx"
LEGACY_UNIT_REPORT_URL = "https://secure.buildinglink.com/v2/mgmt/export/UnitSearch.aspx"
UNIT_REPORT_INITIALIZER_URL = (
    "https://secure.buildinglink.com/v2/mgmt/global/UnitSearchArray.aspx"
)
MANAGER_DASHBOARD_URL = (
    "https://secure.buildinglink.com/manager-dashboard/staff/"
    "?PagePath=%2Fv2%2Fmgmt%2Fdashboard%2Fdashboard.aspx%3F"
)

ROOT = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
DEFAULT_CONFIG = ROOT / "backup-settings.json"


def configure_logging(log_folder: Path) -> Path:
    log_folder.mkdir(parents=True, exist_ok=True)
    log_path = log_folder / f"BuildingLinkBackup_{datetime.now():%Y-%m-%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    return log_path


def load_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"Configuration file not found: {path}")
    with path.open("r", encoding="utf-8-sig") as handle:
        config = json.load(handle)
    validate_config(config)
    return config


def resolve_config_path(value: str, config_path: Path) -> Path:
    path = Path(os.path.expandvars(value)).expanduser()
    if not path.is_absolute():
        path = config_path.parent / path
    return path.resolve()


def validate_config(config: dict[str, Any]) -> None:
    buildings = config.get("buildings")
    if not isinstance(buildings, list) or len(buildings) != 2:
        raise ValueError("The settings file must define exactly two buildings.")
    seen: set[str] = set()
    for building in buildings:
        required = (
            "corporation",
            "facility_id",
            "property_name",
            "occupant_data_folder",
            "unit_data_folder",
        )
        missing = [name for name in required if not str(building.get(name, "")).strip()]
        if missing:
            raise ValueError(f"A building is missing: {', '.join(missing)}")
        corporation = str(building["corporation"]).strip()
        if corporation in seen:
            raise ValueError(f"Duplicate corporation in settings: {corporation}")
        seen.add(corporation)


def credential_backend_name() -> str:
    backend = keyring.get_keyring()
    return f"{type(backend).__module__}.{type(backend).__name__}"


def get_credentials() -> tuple[str, str]:
    try:
        credential = keyring.get_credential(CREDENTIAL_SERVICE, None)
    except KeyringError as exc:
        raise RuntimeError(f"Windows Credential Manager could not be read: {exc}") from exc
    if credential is None or not credential.username or not credential.password:
        raise RuntimeError(
            f"Credential '{CREDENTIAL_SERVICE}' was not found. "
            "Run this program with --setup-credentials first."
        )
    return credential.username, credential.password


def setup_credentials() -> int:
    print(f"Credential backend: {credential_backend_name()}")
    username = input("BuildingLink username: ").strip()
    if not username:
        print("Username cannot be empty.", file=sys.stderr)
        return EXIT_CREDENTIALS
    password = getpass.getpass("BuildingLink password: ")
    if not password:
        print("Password cannot be empty.", file=sys.stderr)
        return EXIT_CREDENTIALS
    try:
        old = keyring.get_credential(CREDENTIAL_SERVICE, None)
        if old is not None and old.username != username:
            try:
                keyring.delete_password(CREDENTIAL_SERVICE, old.username)
            except PasswordDeleteError:
                pass
        keyring.set_password(CREDENTIAL_SERVICE, username, password)
        saved_user, saved_password = get_credentials()
    except (KeyringError, RuntimeError) as exc:
        print(f"Credential could not be saved: {exc}", file=sys.stderr)
        return EXIT_CREDENTIALS
    if saved_user != username or saved_password != password:
        print("Credential verification failed.", file=sys.stderr)
        return EXIT_CREDENTIALS
    print(f"Credential '{CREDENTIAL_SERVICE}' was saved and verified for {username}.")
    return EXIT_SUCCESS


def test_credentials() -> int:
    print(f"Credential backend: {credential_backend_name()}")
    try:
        username, password = get_credentials()
    except RuntimeError as exc:
        print(exc, file=sys.stderr)
        return EXIT_CREDENTIALS
    print(f"Credential found for: {username}")
    print(f"Password available: {'Yes' if password else 'No'}")
    return EXIT_SUCCESS


def delete_credentials() -> int:
    try:
        credential = keyring.get_credential(CREDENTIAL_SERVICE, None)
        if credential is None:
            print(f"Credential '{CREDENTIAL_SERVICE}' is already absent.")
            return EXIT_SUCCESS
        keyring.delete_password(CREDENTIAL_SERVICE, credential.username)
    except PasswordDeleteError:
        print(f"Credential '{CREDENTIAL_SERVICE}' is already absent.")
        return EXIT_SUCCESS
    except KeyringError as exc:
        print(f"Credential could not be deleted: {exc}", file=sys.stderr)
        return EXIT_CREDENTIALS
    print(f"Credential '{CREDENTIAL_SERVICE}' was deleted.")
    return EXIT_SUCCESS


@contextmanager
def single_instance():
    if os.name != "nt":
        yield
        return
    import ctypes

    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel32.CreateMutexW(None, False, f"Local\\{APP_NAME}")
    if not handle:
        raise RuntimeError("Could not create the single-instance lock.")
    try:
        if kernel32.GetLastError() == 183:
            raise RuntimeError("Another BuildingLink backup is already running.")
        yield
    finally:
        kernel32.CloseHandle(handle)


def describe_tabs(context: BrowserContext) -> str:
    values: list[str] = []
    for page in context.pages:
        parsed = urlparse(page.url)
        values.append(f"{parsed.hostname or '(blank)'}{parsed.path or ''}")
    return ", ".join(values)


def allowed_pages(context: BrowserContext) -> list[Page]:
    pages = [page for page in context.pages if not page.is_closed()]
    for page in pages:
        parsed = urlparse(page.url)
        if page.url in {"", "about:blank"} or parsed.scheme == "edge":
            continue
        if parsed.hostname not in {"secure.buildinglink.com", "auth.buildinglink.com"}:
            raise RuntimeError(f"Unexpected browser tab: {parsed.hostname or page.url}")
    return pages


def first_visible(page: Page, selectors: list[str]):
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.is_visible(timeout=750):
                return locator
        except PlaywrightError:
            pass
    return None


def find_page(context: BrowserContext, role: str) -> Page | None:
    for page in reversed(allowed_pages(context)):
        parsed = urlparse(page.url)
        path = parsed.path.rstrip("/").lower()
        try:
            if role == "login" and page.locator("#UserName").is_visible(timeout=500):
                return page
            if role == "properties" and page.get_by_text(
                "Login to this building", exact=False
            ).first.is_visible(timeout=500):
                return page
            if role == "manager" and path.startswith("/manager-dashboard"):
                return page
            if role == "occupant" and path.endswith("download-occupantdata-requests"):
                return page
            if role == "unit" and path == LEGACY_UNIT_REPORT_PATH:
                return page
        except PlaywrightError:
            continue
    return None


def wait_for_page(context: BrowserContext, role: str, timeout_seconds: int) -> Page:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        page = find_page(context, role)
        if page is not None:
            page.bring_to_front()
            return page
        time.sleep(0.25)
    raise RuntimeError(
        f"No {role} page appeared within {timeout_seconds} seconds. "
        f"Open tabs: {describe_tabs(context)}"
    )


def launch_clean_edge(playwright, startup_url: str, headless: bool, action_timeout: int):
    profile = Path(tempfile.mkdtemp(prefix="BuildingLinkBackup-Edge-"))
    context: BrowserContext | None = None
    try:
        context = playwright.chromium.launch_persistent_context(
            str(profile),
            channel="msedge",
            headless=headless,
            accept_downloads=True,
            slow_mo=0 if headless else 200,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-sync",
                "--disable-extensions",
                "--disable-session-crashed-bubble",
                "--hide-crash-restore-bubble",
                "--disable-background-mode",
            ],
            timeout=action_timeout * 1000,
        )
        context.set_default_timeout(action_timeout * 1000)
        context.set_default_navigation_timeout(action_timeout * 1000)
        page = context.pages[0] if context.pages else context.new_page()
        logging.info("Opening BuildingLink in a clean browser session.")
        page.goto(startup_url, wait_until="commit", timeout=action_timeout * 1000)
        if page.url in {"", "about:blank"}:
            raise RuntimeError("Edge remained at about:blank.")
        return context, page, profile
    except Exception:
        if context is not None:
            try:
                context.close()
            except PlaywrightError:
                pass
        shutil.rmtree(profile, ignore_errors=True)
        raise


def close_clean_edge(context: BrowserContext | None, profile: Path | None) -> None:
    if context is not None:
        try:
            context.close()
        except PlaywrightError:
            pass
    if profile is not None:
        for _ in range(8):
            shutil.rmtree(profile, ignore_errors=True)
            if not profile.exists():
                return
            time.sleep(0.25)


def submit_login(page: Page, username: str, password: str, action_timeout: int) -> None:
    logging.info("Login page detected; entering the stored credential.")
    page.locator("#UserName").fill(username)
    page.locator("#Password").fill(password)
    remember = page.locator("#RememberLogin")
    if remember.is_visible(timeout=2_000) and remember.is_checked():
        remember.uncheck(timeout=action_timeout * 1000)
    logging.info("Keep me logged in remains unchecked.")
    page.locator("#LoginButton").click(no_wait_after=True, timeout=action_timeout * 1000)
    page.wait_for_timeout(UI_SETTLE_MS)


def open_building(context: BrowserContext, property_page: Page, property_name: str) -> Page:
    heading = property_page.get_by_text(property_name, exact=True)
    heading.wait_for(state="visible")
    card = heading.locator(
        "xpath=ancestor::div[.//button[contains(normalize-space(.), 'Login to this building')]][1]"
    )
    logging.info("Selecting %s.", property_name)
    card.get_by_role("button", name="Login to this building").click(no_wait_after=True)
    try:
        return wait_for_page(context, "manager", 8)
    except RuntimeError:
        logging.info("Reopening BuildingLink's manager handoff in the tracked tab.")
        property_page.wait_for_timeout(2_000)
        property_page.goto(MANAGER_DASHBOARD_URL, wait_until="commit")
        return property_page


def login_and_open_building(
    context: BrowserContext,
    startup_page: Page,
    username: str,
    password: str,
    property_name: str,
    login_timeout: int,
    action_timeout: int,
    progress=None,
) -> Page:
    deadline = time.monotonic() + login_timeout
    login_submitted = False
    page = startup_page
    while time.monotonic() < deadline:
        manager = find_page(context, "manager")
        if manager is not None:
            return manager
        properties = find_page(context, "properties")
        if properties is not None:
            if progress is not None:
                progress("selecting_building")
            return open_building(context, properties, property_name)
        login = find_page(context, "login")
        if login is not None and not login_submitted:
            if progress is not None:
                progress("login_form")
            submit_login(login, username, password, action_timeout)
            login_submitted = True
            if progress is not None:
                progress("login_submitted")
            page = login
            continue
        for candidate in reversed(allowed_pages(context)):
            if urlparse(candidate.url).path.rstrip("/").lower() == "/unauthorized":
                retry = candidate.get_by_role("button", name="Try again", exact=True)
                if retry.is_visible(timeout=1_000):
                    retry.click(no_wait_after=True)
                else:
                    candidate.goto(PORTAL_URL, wait_until="commit")
                break
        page.wait_for_timeout(500)
    raise RuntimeError(
        f"Login did not reach a BuildingLink dashboard. Open tabs: {describe_tabs(context)}"
    )


def verify_or_switch_building(
    page: Page, property_name: str, facility_id: str, action_timeout: int
) -> None:
    property_id = page.get_by_test_id("property-id")
    property_id.wait_for(state="visible", timeout=action_timeout * 1000)
    if facility_id in property_id.inner_text() and page.get_by_text(
        property_name, exact=True
    ).is_visible(timeout=2_000):
        logging.info("Active corporation verified: %s (%s).", property_name, facility_id)
        return
    logging.info("Switching to %s (%s).", property_name, facility_id)
    dropdown = page.locator("[data-testid='properties-dropdown']:visible").last
    dropdown.wait_for(state="visible")
    dropdown.click()
    page.wait_for_timeout(UI_SETTLE_MS)
    option = page.get_by_test_id("staff-properties-dropdown-on-switch-click-div").filter(
        has_text=property_name
    ).first
    option.wait_for(state="visible")
    option.click(no_wait_after=True)
    deadline = time.monotonic() + action_timeout
    while time.monotonic() < deadline:
        try:
            if facility_id in property_id.inner_text(timeout=1_000):
                logging.info("Corporation switch verified.")
                return
        except PlaywrightError:
            pass
        time.sleep(0.25)
    raise RuntimeError(f"Corporation switch to {property_name} did not complete.")


def navigate_to_download_data(page: Page) -> None:
    logging.info("Opening Reports/Data.")
    page.get_by_text("Reports/Data", exact=True).click(no_wait_after=True)
    page.wait_for_timeout(UI_SETTLE_MS)
    logging.info("Opening Download data.")
    link = page.get_by_text("Download data", exact=True)
    link.wait_for(state="visible")
    link.click(no_wait_after=True)
    page.wait_for_timeout(UI_SETTLE_MS)


def open_occupant_report(
    context: BrowserContext,
    manager: Page,
    property_name: str,
    headless: bool = False,
) -> Page:
    navigate_to_download_data(manager)
    logging.info("Opening Occupant data.")
    if headless:
        # In headless mode the visual tab strip can retain Unit Profile as its
        # active target even when the Occupant label is clicked. The report URL
        # is stable and keeps the same authenticated corporation context.
        manager.goto(
            OCCUPANT_REPORT_URL,
            wait_until="domcontentloaded",
            timeout=DEFAULT_ACTION_TIMEOUT_SECONDS * 1000,
        )
        manager.get_by_text(property_name, exact=True).wait_for(state="visible")
        return manager
    link = manager.get_by_text("Occupant data", exact=True)
    link.wait_for(state="visible")
    link.click(no_wait_after=True)
    page = wait_for_page(context, "occupant", DEFAULT_ACTION_TIMEOUT_SECONDS)
    page.get_by_text(property_name, exact=True).wait_for(state="visible")
    return page


def download_occupant(page: Page, timeout_seconds: int):
    for test_id in (
        "staff-download-occupant-tenant-checkbox",
        "staff-download-occupant-subtenant-checkbox",
    ):
        control = page.get_by_test_id(test_id)
        control.wait_for(state="visible")
        control.click()
    with page.expect_download(timeout=timeout_seconds * 1000) as info:
        page.get_by_role("button", name="Export Excel").click()
    return info.value


def find_unit_form(context: BrowserContext, timeout_seconds: int):
    """Find the legacy Unit Profile form in a top document or child frame."""
    legacy_page = wait_for_page(context, "unit", timeout_seconds)
    logging.info("Legacy Unit Profile page committed: %s", LEGACY_UNIT_REPORT_PATH)
    deadline = time.monotonic() + timeout_seconds
    last_frames = []

    # BuildingLink initially commits the legacy report page and then inserts the
    # actual form into one of its frames. Poll every frame briefly instead of
    # spending a large part of the overall timeout waiting on the wrong frame.
    while time.monotonic() < deadline:
        all_frames = list(legacy_page.frames)
        matching_frames = [
            frame
            for frame in all_frames
            if urlparse(frame.url).path.lower() == LEGACY_UNIT_REPORT_PATH
        ]
        other_frames = [frame for frame in all_frames if frame not in matching_frames]
        frames = matching_frames + other_frames
        last_frames = frames

        for frame in frames:
            try:
                control = frame.locator("#ctl00_MainContentPlaceHolder_chk_altmailing")
                control.wait_for(state="attached", timeout=500)
                logging.info(
                    "Unit Profile form detected in %s.",
                    urlparse(frame.url).path or "the current document",
                )
                legacy_page.bring_to_front()
                return legacy_page, frame
            except PlaywrightError:
                continue

        time.sleep(0.1)

    frame_urls = [urlparse(frame.url).path or "(blank)" for frame in last_frames]
    raise RuntimeError(
        f"Legacy Unit Profile page opened, but its form did not appear within "
        f"{timeout_seconds} seconds. Frames: {', '.join(frame_urls)}"
    )


def open_unit_report(
    context: BrowserContext,
    page: Page,
    property_name: str,
    timeout_seconds: int,
    headless: bool = False,
):
    logging.info("Opening Unit profile data.")
    if headless:
        # The visible BuildingLink workflow opens a legacy report shell through
        # UnitSearchArray.aspx before it loads UnitSearch.aspx. In headless Edge,
        # that popup can remain about:blank, so reproduce the same server-side
        # initialization in the authenticated context and keep one tracked tab.
        logging.info("Initializing the legacy Unit Profile report for headless mode.")
        for _ in range(2):
            response = context.request.get(
                UNIT_REPORT_INITIALIZER_URL,
                headers={"Referer": page.url},
                timeout=timeout_seconds * 1000,
            )
            if not response.ok:
                raise RuntimeError(
                    f"Unit Profile initializer returned HTTP {response.status}."
                )
        form_response = context.request.get(
            LEGACY_UNIT_REPORT_URL,
            headers={"Referer": UNIT_REPORT_INITIALIZER_URL},
            timeout=timeout_seconds * 1000,
        )
        if not form_response.ok:
            raise RuntimeError(
                f"Unit Profile form returned HTTP {form_response.status}."
            )
        parser = UnitStateParser()
        parser.feed(form_response.text())
        missing = [
            name
            for name in ("__EVENTVALIDATION", "__PAGESTATEID")
            if name not in parser.values
        ]
        if missing:
            raise RuntimeError(
                "Unit Profile form did not contain " + ", ".join(missing) + "."
            )
        logging.info("Legacy Unit Profile form initialized without opening a popup.")
        return page, parser.values

    # The Occupant Data page already contains the report tabs. Clicking its Unit
    # tab is more reliable than starting a second direct navigation while the
    # legacy BuildingLink report shell is still active.
    tab = page.get_by_role("link", name="Unit profile data", exact=True)
    logging.info("Waiting for the Unit profile data tab.")
    tab.wait_for(state="visible", timeout=timeout_seconds * 1000)
    logging.info("Unit profile data tab is visible; selecting it.")
    tab.click(no_wait_after=True, timeout=timeout_seconds * 1000)
    logging.info("Waiting for the legacy Unit Profile form.")
    return find_unit_form(context, timeout_seconds)


def unit_checkbox(surface, label: str, selector: str):
    """Prefer the visible label and retain the stable legacy ID as a fallback."""
    labelled = surface.get_by_label(label, exact=True)
    if labelled.count():
        return labelled.first
    return surface.locator(selector).first


def ensure_checkbox(surface, label: str, selector: str, checked: bool) -> None:
    box = unit_checkbox(surface, label, selector)
    box.wait_for(state="attached")
    if checked and not box.is_checked():
        box.check()
    elif not checked and box.is_checked():
        box.uncheck()


class ResponseDownload:
    """Small adapter that lets an API response use the normal download saver."""

    def __init__(self, suggested_filename: str, content: bytes):
        self.suggested_filename = suggested_filename
        self._content = content

    def save_as(self, path: str | Path) -> None:
        Path(path).write_bytes(self._content)


class UnitStateParser(HTMLParser):
    """Extract named input values from BuildingLink's legacy ASP.NET form."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.values: dict[str, str] = {}

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() != "input":
            return
        values = dict(attrs)
        name = values.get("name")
        if name:
            self.values[name] = values.get("value", "")


def download_unit(
    page: Page,
    surface,
    timeout_seconds: int,
    context: BrowserContext | None = None,
    headless: bool = False,
):
    if headless:
        if context is None:
            raise RuntimeError("Headless Unit export requires its browser context.")
        if not isinstance(surface, dict):
            raise RuntimeError("Headless Unit export did not receive initialized form state.")
        logging.info("Submitting the initialized Unit Profile form in headless mode.")
        payload = {
            "__EVENTTARGET": "ctl00$MainContentPlaceHolder$ExportButton",
            "__EVENTVALIDATION": surface["__EVENTVALIDATION"],
            "__PAGESTATEID": surface["__PAGESTATEID"],
            "ctl00$MainContentPlaceHolder$chk_altmailing": "on",
            "ctl00$MainContentPlaceHolder$chk_custom": "on",
            "ctl00$MainContentPlaceHolder$rb_insert_building_address": "0",
            "ctl00$MainContentPlaceHolder$rb_include_deactivated_units": "0",
        }
        response = context.request.post(
            LEGACY_UNIT_REPORT_URL,
            form=payload,
            headers={"Referer": LEGACY_UNIT_REPORT_URL},
            timeout=timeout_seconds * 1000,
        )
        disposition = response.headers.get("content-disposition", "")
        if not response.ok or "attachment" not in disposition.lower():
            raise RuntimeError(
                f"Unit Profile export returned HTTP {response.status} without an attachment."
            )
        match = re.search(r"filename\s*=\s*[\"']?([^\"';]+)", disposition, re.I)
        filename = match.group(1).strip() if match else "UnitProfile.xls"
        return ResponseDownload(filename, response.body())

    ensure_checkbox(
        surface,
        "Alt Mailing Address",
        "#ctl00_MainContentPlaceHolder_chk_altmailing",
        True,
    )
    ensure_checkbox(
        surface,
        "Custom Fields",
        "#ctl00_MainContentPlaceHolder_chk_custom",
        True,
    )
    optional_off = (
        ("Occupant Contact Info", "#ctl00_MainContentPlaceHolder_chk_occupant"),
        ("Advanced Fields", "#ctl00_MainContentPlaceHolder_chk_advanced"),
    )
    for label, selector in optional_off:
        if surface.get_by_label(label, exact=True).count() or surface.locator(selector).count():
            ensure_checkbox(surface, label, selector, False)
    export = surface.locator("#ctl00_MainContentPlaceHolder_ExportButton")
    if not export.count():
        export = surface.get_by_role("link", name="Export To Excel", exact=True)
    with page.expect_download(timeout=timeout_seconds * 1000) as info:
        export.first.click()
    return info.value


def workbook_text(path: Path) -> str:
    import xml.etree.ElementTree as element_tree

    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    parts: list[str] = []
    with zipfile.ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = element_tree.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{namespace}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{namespace}t")))
        parts.extend(shared)
        for name in archive.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            root = element_tree.fromstring(archive.read(name))
            for cell in root.iter(f"{namespace}c"):
                value = cell.find(f"{namespace}v")
                if value is None or value.text is None:
                    continue
                if cell.attrib.get("t") == "s":
                    try:
                        parts.append(shared[int(value.text)])
                    except (ValueError, IndexError):
                        pass
                else:
                    parts.append(value.text)
    return "\n".join(parts)


def validate_workbook(path: Path, report_type: str, corporation: str) -> None:
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError("Downloaded workbook is missing or empty.")
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise RuntimeError(f"Unexpected downloaded file type: {path.suffix}")
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        if not zipfile.is_zipfile(path):
            raise RuntimeError("Downloaded file is not a valid Excel workbook.")
        with zipfile.ZipFile(path) as archive:
            required = {"xl/workbook.xml", "[Content_Types].xml"}
            if not required.issubset(set(archive.namelist())):
                raise RuntimeError("Downloaded archive is missing Excel workbook content.")
        text = workbook_text(path).lower()
        if report_type == "Occupant-Data":
            if "occupants data" not in text and "email" not in text:
                raise RuntimeError("Workbook does not appear to contain occupant data.")
            if corporation.lower() not in text:
                raise RuntimeError(f"Workbook does not contain {corporation}.")
        elif not any(term in text for term in ("unit id", "unit", "mailing")):
            raise RuntimeError("Workbook does not appear to contain unit profile data.")


def save_staged_download(
    download,
    staging: Path,
    corporation: str,
    report_type: str,
) -> Path:
    suffix = Path(download.suggested_filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        suffix = ".xlsx"
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = f"{corporation.replace(' ', '-')}_{report_type}_{stamp}{suffix}"
    local = staging / filename
    download.save_as(str(local))
    validate_workbook(local, report_type, corporation)
    logging.info("Downloaded and validated locally: %s", local.name)
    return local


def copy_verified_workbook(
    local: Path, destination_folder: Path, report_type: str, corporation: str
) -> Path:
    destination_folder.mkdir(parents=True, exist_ok=True)
    destination = destination_folder / local.name
    if destination.exists():
        destination = destination_folder / f"{local.stem}_{time.time_ns()}{local.suffix}"
    shutil.copy2(local, destination)
    validate_workbook(destination, report_type, corporation)
    logging.info("Saved and verified: %s", destination)
    return destination


def run_building(
    playwright,
    config: dict[str, Any],
    config_path: Path,
    building: dict[str, Any],
    username: str,
    password: str,
) -> list[Path]:
    corporation = str(building["corporation"]).strip()
    property_name = str(building["property_name"]).strip()
    facility_id = str(building["facility_id"]).strip()
    occupant_folder = resolve_config_path(str(building["occupant_data_folder"]), config_path)
    unit_folder = resolve_config_path(str(building["unit_data_folder"]), config_path)
    headless = bool(config.get("headless", True))
    action_timeout = int(config.get("action_timeout_seconds", DEFAULT_ACTION_TIMEOUT_SECONDS))
    login_timeout = int(config.get("login_timeout_seconds", DEFAULT_LOGIN_TIMEOUT_SECONDS))
    download_timeout = int(config.get("download_timeout_seconds", DEFAULT_DOWNLOAD_TIMEOUT_SECONDS))
    max_attempts = int(config.get("max_attempts", DEFAULT_MAX_ATTEMPTS))

    last_error: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        context: BrowserContext | None = None
        profile: Path | None = None
        try:
            logging.info("%s attempt %s of %s.", corporation, attempt, max_attempts)
            context, page, profile = launch_clean_edge(
                playwright, PORTAL_URL, headless, action_timeout
            )
            manager = login_and_open_building(
                context,
                page,
                username,
                password,
                property_name,
                login_timeout,
                action_timeout,
            )
            verify_or_switch_building(manager, property_name, facility_id, action_timeout)

            with tempfile.TemporaryDirectory(prefix=f"{corporation.replace(' ', '-')}-") as tmp:
                staging = Path(tmp)
                occupant_page = open_occupant_report(
                    context, manager, property_name, headless=headless
                )
                occupant = download_occupant(occupant_page, download_timeout)
                staged_occupant = save_staged_download(
                    occupant, staging, corporation, "Occupant-Data"
                )
                unit_page, unit_surface = open_unit_report(
                    context,
                    occupant_page,
                    property_name,
                    download_timeout,
                    headless=headless,
                )
                unit = download_unit(
                    unit_page,
                    unit_surface,
                    download_timeout,
                    context=context,
                    headless=headless,
                )
                staged_unit = save_staged_download(
                    unit, staging, corporation, "Unit-Data"
                )
                # Publish both reports only after both passed local validation.
                occupant_path = copy_verified_workbook(
                    staged_occupant, occupant_folder, "Occupant-Data", corporation
                )
                unit_path = copy_verified_workbook(
                    staged_unit, unit_folder, "Unit-Data", corporation
                )
            logging.info("%s completed successfully.", corporation)
            return [occupant_path, unit_path]
        except Exception as exc:
            last_error = exc
            logging.exception("%s attempt %s failed: %s", corporation, attempt, exc)
        finally:
            close_clean_edge(context, profile)
            logging.info("Closed the clean browser session for %s.", corporation)
        if attempt < max_attempts:
            time.sleep(2)
    raise RuntimeError(f"{corporation} failed after {max_attempts} attempts: {last_error}")


def run_backup(config: dict[str, Any], config_path: Path) -> int:
    username, password = get_credentials()
    log_folder = resolve_config_path(str(config.get("log_folder", "logs")), config_path)
    log_path = configure_logging(log_folder)
    logging.info("Starting BuildingLink daily backup.")
    logging.info("Credential found for %s using %s.", username, credential_backend_name())
    completed: list[Path] = []
    try:
        with single_instance(), sync_playwright() as playwright:
            for building in config["buildings"]:
                completed.extend(
                    run_building(playwright, config, config_path, building, username, password)
                )
    except Exception as exc:
        logging.exception("BACKUP FAILED: %s", exc)
        logging.info("Log file: %s", log_path)
        return EXIT_AUTOMATION
    logging.info("BACKUP SUCCESS: all four files were saved and verified.")
    for path in completed:
        logging.info("Result: %s", path)
    logging.info("Log file: %s", log_path)
    return EXIT_SUCCESS


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    actions = parser.add_mutually_exclusive_group()
    actions.add_argument("--setup-credentials", action="store_true")
    actions.add_argument("--test-credentials", action="store_true")
    actions.add_argument("--delete-credentials", action="store_true")
    actions.add_argument("--validate-config", action="store_true")
    parser.add_argument("--headed", action="store_true", help="Show Edge for troubleshooting.")
    args = parser.parse_args()

    if args.setup_credentials:
        return setup_credentials()
    if args.test_credentials:
        return test_credentials()
    if args.delete_credentials:
        return delete_credentials()

    config_path = args.config.resolve()
    try:
        config = load_config(config_path)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"Configuration error: {exc}", file=sys.stderr)
        return EXIT_CONFIG
    if args.validate_config:
        print(f"Configuration is valid: {config_path}")
        return EXIT_SUCCESS
    if args.headed:
        config["headless"] = False
    try:
        return run_backup(config, config_path)
    except RuntimeError as exc:
        print(exc, file=sys.stderr)
        return EXIT_CREDENTIALS
    except KeyboardInterrupt:
        print("Backup cancelled by the user.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit("Public showcase: live worker execution is disabled.")
    raise SystemExit(main())
