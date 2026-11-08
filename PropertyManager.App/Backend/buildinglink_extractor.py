from __future__ import annotations

import hashlib
import re
import shutil
import tempfile
import time
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse
from booking_identity import BookingIdentityError, identify_requested_for

from playwright.sync_api import BrowserContext, Error as PlaywrightError, Frame, Page, sync_playwright


PORTAL_URL = "https://secure.buildinglink.com/"
RESERVATION_URL = "https://secure.buildinglink.com/v2/mgmt/amenities/listview.aspx"
SHARED_PROPERTY = "Example Condominium - DEMO 1001"
MANAGER_DASHBOARD_URL = "https://secure.buildinglink.com/manager-dashboard/staff/?PagePath=%2Fv2%2Fmgmt%2Fdashboard%2Fdashboard.aspx%3F"


@dataclass(frozen=True)
class AmenityBooking:
    reservation_id: str
    start: datetime
    end: datetime
    amenity: str
    requested_for: str
    description: str
    status: str
    corporation: str = ""
    unit: str = ""


def _parse_datetime(value: str) -> datetime:
    value = " ".join(value.split())
    for fmt in (
        "%d %b %Y %H:%M",
        "%d %B %Y %H:%M",
        "%m/%d/%y %I:%M %p",
        "%m/%d/%Y %I:%M %p",
    ):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unrecognized BuildingLink date/time: {value!r}")


def parse_reservation_time(value: str) -> tuple[datetime, datetime]:
    parts = re.split(r"\s+-\s+", " ".join(value.split()), maxsplit=1)
    if len(parts) != 2:
        raise ValueError(f"Unrecognized reservation time range: {value!r}")
    start = _parse_datetime(parts[0])
    for time_format in ("%H:%M", "%I:%M %p"):
        try:
            end_time = datetime.strptime(parts[1], time_format).time()
            return start, datetime.combine(start.date(), end_time)
        except ValueError:
            pass
    return start, _parse_datetime(parts[1])


def _booking_id(values: list[str]) -> str:
    canonical = "|".join(" ".join(value.casefold().split()) for value in values)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16]


def parse_booking_row(cells: list[str]) -> AmenityBooking | None:
    if len(cells) < 5 or cells[0].strip().casefold() == "reservation time":
        return None
    values = [cell.strip() for cell in cells[:5]]
    values[2] = values[2].replace("\ufffdE", "·")
    start, end = parse_reservation_time(values[0])
    return AmenityBooking(_booking_id(values), start, end, values[1], values[2], values[3], values[4])


def _allowed_pages(context: BrowserContext) -> list[Page]:
    pages = [page for page in context.pages if not page.is_closed()]
    for page in pages:
        parsed = urlparse(page.url)
        if page.url in {"", "about:blank"} or parsed.scheme == "edge":
            continue
        if parsed.hostname not in {"secure.buildinglink.com", "auth.buildinglink.com"}:
            raise RuntimeError(f"Unexpected browser tab during login: {parsed.hostname or page.url}")
    return pages


def _find_page(context: BrowserContext, role: str) -> Page | None:
    for page in reversed(_allowed_pages(context)):
        path = urlparse(page.url).path.rstrip("/").casefold()
        try:
            if role == "login" and page.locator("#UserName").is_visible(timeout=500):
                return page
            if role == "properties" and page.get_by_text("Login to this building", exact=False).first.is_visible(timeout=500):
                return page
            if role == "manager" and (path.startswith("/manager-dashboard") or path == "/v2/default.aspx"):
                return page
        except PlaywrightError:
            continue
    return None


def _wait_for_page(context: BrowserContext, role: str, timeout_seconds: int) -> Page:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        page = _find_page(context, role)
        if page is not None:
            return page
        time.sleep(0.25)
    raise RuntimeError(f"BuildingLink login did not reach the {role} page within {timeout_seconds} seconds.")


def _login(context: BrowserContext, page: Page, username: str, password: str, property_name: str, timeout: int, progress=None) -> Page:
    deadline = time.monotonic() + timeout
    submitted = 0
    last_report = 0.0
    while time.monotonic() < deadline:
        if progress and time.monotonic()-last_report >= 3:
            progress("Waiting for BuildingLink sign-in and property selection")
            last_report=time.monotonic()
        manager = _find_page(context, "manager")
        if manager is not None:
            return manager
        properties = _find_page(context, "properties")
        if properties is not None:
            heading = properties.get_by_text(property_name, exact=True)
            heading.wait_for(state="visible")
            card = heading.locator("xpath=ancestor::div[.//button[contains(normalize-space(.), 'Login to this building')]][1]")
            card.get_by_role("button", name="Login to this building").click(no_wait_after=True)
            try:
                return _wait_for_page(context, "manager", 8)
            except RuntimeError:
                try:
                    properties.goto(MANAGER_DASHBOARD_URL, wait_until="commit")
                except PlaywrightError:
                    pass
                page = properties
                continue
        login = _find_page(context, "login")
        if login is not None and submitted < 3:
            login.locator("#UserName").fill(username)
            login.locator("#Password").fill(password)
            remember = login.locator("#RememberLogin")
            if remember.is_visible(timeout=2_000) and remember.is_checked():
                remember.uncheck()
            login.locator("#LoginButton").click(no_wait_after=True)
            submitted += 1
            page = login
        page.wait_for_timeout(500)
    raise RuntimeError("BuildingLink login did not reach a manager page.")


def _set_date_filters(page: Page, start_date: date, end_date: date) -> None:
    _pause(page, 5_000)
    date_inputs = page.locator("input:visible")
    matches = []
    for index in range(date_inputs.count()):
        item = date_inputs.nth(index)
        value = item.input_value()
        placeholder = item.get_attribute("placeholder") or ""
        if re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}", value) or "date" in placeholder.casefold():
            matches.append(item)
    if len(matches) < 2:
        metadata = []
        for index in range(date_inputs.count()):
            item = date_inputs.nth(index)
            if (item.get_attribute("type") or "").casefold() == "password":
                continue
            metadata.append({
                "type": item.get_attribute("type"),
                "name": item.get_attribute("name"),
                "placeholder": item.get_attribute("placeholder"),
                "aria_label": item.get_attribute("aria-label"),
            })
        frame_urls = [frame.url for frame in page.frames]
        snippet = " ".join(page.locator("body").inner_text().split())[:500]
        raise RuntimeError(
            f"Could not locate BuildingLink's date filters. Visible inputs: {metadata}; "
            f"frames: {frame_urls}; page text: {snippet!r}"
        )
    language = str(page.evaluate("() => navigator.language || document.documentElement.lang || ''")).casefold()
    input_type = (matches[0].get_attribute("type") or "text").casefold()
    placeholder = (matches[0].get_attribute("placeholder") or "").casefold()

    def formatted(value: date) -> str:
        if input_type == "date":
            return value.strftime("%Y-%m-%d")
        if "mm/dd" in placeholder or language in {"en-us", "en_us"}:
            return value.strftime("%m/%d/%Y")
        return value.strftime("%d/%m/%Y")

    for item, value in zip(matches[:2], (formatted(start_date), formatted(end_date))):
        item.fill(value)
        item.dispatch_event("change")
    matches[1].press("Enter")
    _pause(page, 2_500)
    if page.get_by_text("Invalid date format", exact=False).first.is_visible(timeout=1_000):
        raise RuntimeError(
            f"BuildingLink rejected the date filter format for browser language {language or '(unknown)'}."
        )


def _pause(surface, milliseconds):
    (surface.page if isinstance(surface, Frame) else surface).wait_for_timeout(milliseconds)


def _reservation_surface(page, timeout=60):
    """The new amenity list is embedded inside the legacy manager shell."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for frame in page.frames:
            if urlparse(frame.url).path.rstrip('/').casefold().startswith('/amenities/staff'):
                return frame
        if page.get_by_text('Reservation time', exact=True).first.is_visible(timeout=250):
            return page
        page.wait_for_timeout(100)
    raise RuntimeError('BuildingLink did not load the amenity reservation list.')


def _detail_form(context, timeout):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for candidate in reversed(_allowed_pages(context)):
            for frame in candidate.frames:
                if urlparse(frame.url).path.casefold().endswith('/amenities/editreservation.aspx'):
                    form = frame.locator('#aspnetForm')
                    if form.is_visible(timeout=250):
                        return candidate, form
        time.sleep(0.1)
    raise BookingIdentityError('The selected reservation did not open its details.')


def _read_booking_identity(manager, cells, booking, start_date, end_date, timeout, report):
    surface = _reservation_surface(manager, timeout)
    wanted = [' '.join(value.split()) for value in cells[:5]]
    candidates = surface.locator('table tr')
    matching = []
    for i in range(candidates.count()):
        row = candidates.nth(i)
        values = row.locator('td').all_inner_texts()
        if [' '.join(value.split()) for value in values[:5]] == wanted:
            matching.append(row)
    if len(matching) != 1:
        raise BookingIdentityError('The reservation row is missing or duplicated. Refresh before retrying.')
    list_url = manager.url
    original_pages = list(manager.context.pages)
    detail_page = None
    try:
        # Opening the row only; never click Edit, status controls, Save or Send Email.
        matching[0].click()
        detail_page, form = _detail_form(manager.context, timeout)
        header = detail_page.locator('#vue-application-navigation-header')
        header.wait_for(state='visible', timeout=timeout*1000)
        header_text = header.inner_text()
        # Local bookings have no "from" note. Only use a positively verified
        # DEMO 1001 context; never infer corporation from unit number or amenity.
        if not re.search(r'\bDEMO\s+1001\b', header_text, re.I) or re.search(r'\bDEMO\s+2002\b', header_text, re.I):
            raise BookingIdentityError('The shared reservation calendar is not in the verified DEMO 1001 context.')
        corporation, unit, requested = identify_requested_for(form.inner_text(), SHARED_PROPERTY, booking.requested_for)
        return replace(booking, corporation=corporation, unit=unit, requested_for=requested)
    finally:
        if detail_page is not None and detail_page not in original_pages:
            detail_page.close()
        else:
            manager.goto(list_url, wait_until='domcontentloaded')
            # Reapply dates after returning; never rely on browser filter persistence.
            _set_date_filters(_reservation_surface(manager, timeout), start_date, end_date)


def _read_rows(page: Page) -> list[list[str]]:
    table_rows = page.locator("table tr")
    rows: list[list[str]] = []
    for index in range(table_rows.count()):
        cells = table_rows.nth(index).locator("th,td")
        values = [" ".join(cells.nth(col).inner_text().split()) for col in range(cells.count())]
        if len(values) >= 5 and values[0]:
            rows.append(values)
    if rows:
        return rows
    return page.evaluate("""
      () => {
        const h = [...document.querySelectorAll('*')].find(e => e.children.length === 0 && e.textContent.trim() === 'Reservation time');
        if (!h) return [];
        const root = h.parentElement.parentElement;
        return [...root.children].map(r => [...r.children].map(c => c.innerText.trim())).filter(r => r.length >= 5);
      }
    """)


def _open_reservations(page: Page, url: str) -> None:
    page.wait_for_timeout(2_000)
    last_error: Exception | None = None
    for _ in range(3):
        try:
            page.goto(url, wait_until="domcontentloaded")
            current_path = urlparse(page.url).path.casefold()
            if current_path == urlparse(url).path.casefold() or current_path.startswith("/amenities/staff"):
                return
        except PlaywrightError as exc:
            last_error = exc
        page.wait_for_timeout(2_000)
    raise RuntimeError(f"Could not open Amenity Reservations after login: {last_error or page.url}")


def extract_bookings(username: str, password: str, start_date: date, end_date: date, amenities: set[str], statuses: set[str], property_name: str, url: str = RESERVATION_URL, headless: bool = True, timeout: int = 60, progress=None) -> list[AmenityBooking]:
    def report(message):
        if progress is not None: progress(message)
    report("Starting Microsoft Edge")
    profile = Path(tempfile.mkdtemp(prefix="OperationsReminder-Edge-"))
    context: BrowserContext | None = None
    try:
        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                str(profile), channel="msedge", headless=headless,
                args=["--no-first-run", "--disable-sync", "--disable-extensions", "--disable-background-mode"],
                timeout=timeout * 1000,
            )
            context.set_default_timeout(timeout * 1000)
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(PORTAL_URL, wait_until="commit")
            report("Signing in to BuildingLink")
            manager = _login(context, page, username, password, SHARED_PROPERTY, timeout, progress=report)
            report("Opening amenity reservations")
            _open_reservations(manager, url)
            report("Applying booking dates")
            surface = _reservation_surface(manager, timeout)
            report("Resetting calendar filters to include all amenities")
            reset = surface.get_by_role("button", name="Reset filters", exact=True)
            if reset.is_visible(timeout=1_000):
                reset.click()
            _set_date_filters(surface, start_date, end_date)
            report("Waiting for booking results")
            surface.wait_for_function("""() => {
              const text=document.body.innerText;
              const loading=[...document.querySelectorAll('[aria-busy=true], [role=progressbar]')].some(e=>e.getClientRects().length);
              if(loading)return false;
              return /no (?:matching )?(?:reservations|bookings|records|results)(?: (?:were )?found| available| to display)?/i.test(text) || /\\b0\\s*(?:-|to)?\\s*0?\\s+of\\s+0\\b/i.test(text) || [...document.querySelectorAll('table tr')].some(r=>r.querySelectorAll('td').length>=5) || [...document.querySelectorAll('*')].some(e=>e.children.length===0 && e.textContent.trim()==='Reservation time');
            }""", timeout=timeout*1000)
            report("Reading booking rows")
            bookings: list[AmenityBooking] = []
            raw_rows = _read_rows(surface)
            parse_failures: list[list[str]] = []
            parsed_count = 0
            for cells in raw_rows:
                report("Reading booking row "+str(parsed_count+1))
                try:
                    booking = parse_booking_row(cells)
                except ValueError:
                    parse_failures.append(cells)
                    continue
                if booking is None:
                    continue
                parsed_count += 1
                if amenities and booking.amenity.casefold() not in amenities:
                    continue
                if statuses and booking.status.casefold() not in statuses:
                    continue
                if booking.end.date() < start_date or booking.start.date() > end_date:
                    continue
                report("Checking reservation details: " + str(len(bookings)+1))
                booking = _read_booking_identity(manager, cells, booking, start_date, end_date, timeout, report)
                bookings.append(booking)
            page_text = " ".join(_reservation_surface(manager, timeout).locator("body").inner_text().split())
            reservation_count = re.search(r"\b(\d+)\s+of\s+(\d+)\b", page_text, re.IGNORECASE)
            if reservation_count and int(reservation_count.group(2)) > 0 and parsed_count == 0:
                raise RuntimeError(
                    "BuildingLink displayed reservations but none could be parsed. "
                    f"Raw rows: {raw_rows[:5]}; parse failures: {parse_failures[:5]}; "
                    f"page text: {page_text[:1200]!r}"
                )
            empty_confirmed = re.search(r"\bno (?:matching )?(?:reservations|bookings|records|results)(?: (?:were )?found| available| to display)?|\b0\s*(?:-|to)?\s*0?\s+of\s+0\b", page_text, re.I)
            if parsed_count == 0 and not empty_confirmed:
                raise RuntimeError("The reservation page did not confirm an empty result. Previous Excel bookings were preserved.")
            if parse_failures:
                raise RuntimeError("Some reservation rows could not be read. Previous Excel bookings were preserved.")
            if reservation_count and int(reservation_count.group(2)) > parsed_count:
                raise RuntimeError("The reservation list contains more records than were read. Narrow the date range before retrying.")
            report("Bookings retrieved: "+str(len(bookings)))
            return sorted(bookings, key=lambda item: (item.start, item.amenity.casefold()))
    finally:
        if context is not None:
            try:
                context.close()
            except Exception:
                # Cleanup must not replace a validated result or the original extraction error.
                pass
        if profile.resolve().parent != Path(tempfile.gettempdir()).resolve() or not profile.name.startswith("OperationsReminder-Edge-"):
            raise RuntimeError("Refusing to clean a browser profile outside the temporary profile directory.")
        for _ in range(8):
            shutil.rmtree(profile, ignore_errors=True)
            if not profile.exists():
                break
            time.sleep(0.25)
