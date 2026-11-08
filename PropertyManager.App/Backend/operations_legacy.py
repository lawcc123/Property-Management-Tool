from __future__ import annotations

import argparse
import html
import json
import os
import re
import smtplib
import sys
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any
from urllib.parse import quote

from openpyxl import load_workbook

from buildinglink_extractor import AmenityBooking, extract_bookings


OPEN_STATUSES = {"", "open", "pending", "waiting", "in progress", "follow up", "follow-up", "requested", "approved"}
CLOSED_STATUSES = {"close", "closed", "done", "completed", "cancelled", "canceled", "resolved"}
DATE_HEADERS = {"reported date", "report date"}
AMENITY_SHEET = "Amenity Booking"
AMENITY_HEADERS = [
    "Reservation ID", "Start", "End", "Amenity", "Requested For",
    "Description", "Status", "Extracted At", "Corporation", "Unit",
]


@dataclass(frozen=True)
class RepairCase:
    category: str
    unit: str
    reported_date: date
    description: str
    note: str
    status: str


@dataclass(frozen=True)
class BulkItemRequest:
    property_name: str
    unit: str
    reported_date: date | None
    items: str
    note: str
    status: str


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def parse_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = re.sub(r"^Sept(?=[- ,])", "Sep", str(value).strip(), flags=re.IGNORECASE)
    for fmt in ("%b-%d-%Y", "%b-%d-%y", "%b %d %Y", "%b %d, %Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def is_open_status(status: str) -> bool:
    return status.strip().casefold() not in CLOSED_STATUSES


def _headers(sheet: Any) -> dict[str, int]:
    first_row = next(sheet.iter_rows(min_row=1, max_row=1), ())
    return {text(cell.value).casefold(): index for index, cell in enumerate(first_row, 1) if text(cell.value)}


def read_open_cases(workbook_path: Path) -> list[RepairCase]:
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    cases: list[RepairCase] = []
    try:
        for sheet in workbook.worksheets:
            headers = _headers(sheet)
            if "unit" not in headers or "description" not in headers or "items" in headers:
                continue
            date_col = next((headers[h] for h in DATE_HEADERS if h in headers), None)
            for row in range(2, sheet.max_row + 1):
                reported = parse_date(sheet.cell(row, date_col).value) if date_col else None
                unit = text(sheet.cell(row, headers["unit"]).value)
                description = text(sheet.cell(row, headers["description"]).value)
                status = text(sheet.cell(row, headers.get("status", 0)).value) if headers.get("status") else ""
                if not reported or not (unit or description) or not is_open_status(status):
                    continue
                note = text(sheet.cell(row, headers.get("note", 0)).value) if headers.get("note") else ""
                cases.append(RepairCase(sheet.title, unit or "Unknown unit", reported, description or "(No description)", note, status or "Open"))
    finally:
        workbook.close()
    return cases


def read_bulk_requests(workbook_path: Path) -> list[BulkItemRequest]:
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    requests: list[BulkItemRequest] = []
    try:
        for sheet in workbook.worksheets:
            headers = _headers(sheet)
            if "unit" not in headers or "items" not in headers:
                continue
            date_col = next((headers[h] for h in DATE_HEADERS if h in headers), None)
            for row in range(2, sheet.max_row + 1):
                unit = text(sheet.cell(row, headers["unit"]).value)
                items = text(sheet.cell(row, headers["items"]).value)
                status = text(sheet.cell(row, headers.get("status", 0)).value) if headers.get("status") else ""
                if not (unit or items) or not is_open_status(status):
                    continue
                note_col = headers.get("note") or headers.get("contact")
                requests.append(BulkItemRequest(
                    sheet.title, unit or "Unknown unit",
                    parse_date(sheet.cell(row, date_col).value) if date_col else None,
                    items or "(No item description)",
                    text(sheet.cell(row, note_col).value) if note_col else "",
                    status or "Open",
                ))
    finally:
        workbook.close()
    return requests


def update_amenity_sheet(workbook_path: Path, bookings: list[AmenityBooking], extracted_at: datetime) -> None:
    workbook = load_workbook(workbook_path)
    sheet = workbook[AMENITY_SHEET] if AMENITY_SHEET in workbook.sheetnames else workbook.create_sheet(AMENITY_SHEET)
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)
    sheet.append(AMENITY_HEADERS)
    for booking in bookings:
        sheet.append([
            booking.reservation_id, booking.start, booking.end, booking.amenity,
            booking.requested_for, booking.description, booking.status, extracted_at, booking.corporation, booking.unit,
        ])
    for cell in sheet[1]:
        header_font = copy(cell.font)
        header_font.bold = True
        cell.font = header_font
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row, 2).number_format = "yyyy-mm-dd hh:mm"
        sheet.cell(row, 3).number_format = "yyyy-mm-dd hh:mm"
        sheet.cell(row, 8).number_format = "yyyy-mm-dd hh:mm"
    widths = [20, 19, 19, 20, 30, 45, 14, 19]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[chr(64 + index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:H{max(sheet.max_row, 1)}"
    workbook.save(workbook_path)


def _repair_section(cases: list[RepairCase], max_cases: int) -> list[str]:
    lines = ["REPAIR LOG", f"Total open cases with valid reported dates: {len(cases)}"]
    lines.append("")
    category_order = list(dict.fromkeys(case.category for case in cases))
    for category in category_order:
        all_category_cases = sorted(
            (case for case in cases if case.category == category),
            key=lambda item: (item.reported_date, item.unit.casefold()),
        )
        category_cases = all_category_cases[:max_cases]
        if not category_cases:
            continue
        lines.extend([category, "Reported Date | Unit | Description | Note"])
        for case in category_cases:
            lines.append(
                f"{case.reported_date:%Y-%m-%d} | {case.unit} | {case.description} | {case.note}"
            )
        lines.append("")
        if len(all_category_cases) > max_cases:
            lines.append(
                f"{len(all_category_cases) - max_cases} additional {category} case(s) are not shown."
            )
            lines.append("")
    if not cases:
        lines.append("No open repair cases found.")
    return lines


def _booking_section(bookings: list[AmenityBooking], start_date: date, end_date: date, error: str | None) -> list[str]:
    lines = ["UPCOMING AMENITY BOOKINGS", f"Period: {start_date:%Y-%m-%d} to {end_date:%Y-%m-%d}"]
    if error:
        return lines + [f"WARNING: Booking extraction failed: {error}"]
    if not bookings:
        return lines + ["No matching bookings found."]
    for index, booking in enumerate(bookings, 1):
        time_range = f"{booking.start:%a %b %d, %H:%M} - "
        time_range += f"{booking.end:%H:%M}" if booking.end.date() == booking.start.date() else f"{booking.end:%a %b %d, %H:%M}"
        lines.append(f"{index}. {booking.amenity} | {time_range} | {booking.status}")
        lines.append(f"   {booking.requested_for}" + (f" | {booking.description}" if booking.description else ""))
    return lines


def _bulk_section(requests: list[BulkItemRequest], enabled: bool) -> list[str]:
    lines = ["BULK ITEM REQUESTS"]
    if not enabled:
        return lines + ["This section is not enabled yet."]
    if not requests:
        return lines + ["No open bulk item requests found."]
    for index, request in enumerate(requests, 1):
        reported = request.reported_date.strftime("%Y-%m-%d") if request.reported_date else "Date not entered"
        lines.append(f"{index}. {reported} | {request.property_name} | {request.unit}")
        lines.append(f"   {request.items}" + (f" | Contact/Note: {request.note}" if request.note else ""))
    return lines


def build_email_body(cases: list[RepairCase], bookings: list[AmenityBooking], bulk_requests: list[BulkItemRequest], start_date: date, end_date: date, config: dict[str, Any], extraction_error: str | None = None) -> str:
    information_path = config.get("information_path") or config.get(
        "workbook", "Repair Log, Bulk Item Requests and Amenity Bookings.xlsx"
    )
    sections = [
        f"Property Operations Daily Reminder - {date.today():%Y-%m-%d}",
        f"For more information, please go to {information_path}",
        "",
    ]
    groups = [
        _booking_section(bookings, start_date, end_date, extraction_error),
        _repair_section(cases, int(config.get("max_cases", 20))),
        _bulk_section(bulk_requests, bool(config.get("bulk_items", {}).get("enabled", False))),
    ]
    for index, group in enumerate(groups):
        if index:
            sections.extend(["", "=" * 60, ""])
        sections.extend(group)
    return "\n".join(sections).rstrip() + "\n"


def _html_text(value: Any) -> str:
    return html.escape(text(value)).replace("\n", "<br>")


def _file_link(path_value: str) -> str:
    normalized = path_value.strip()
    if normalized.startswith("\\\\"):
        url = "file://" + quote(normalized.lstrip("\\").replace("\\", "/"), safe="/:@")
    else:
        try:
            url = Path(normalized).resolve().as_uri()
        except ValueError:
            url = "file:///" + quote(normalized.replace("\\", "/"), safe="/:@")
    return (
        f'<a href="{html.escape(url, quote=True)}" '
        f'style="color:#8b3f65; text-decoration:underline;">{_html_text(normalized)}</a>'
    )


def _html_table(headers: list[str], rows: list[list[Any]], widths: list[str] | None = None) -> str:
    header_cells = []
    for index, label in enumerate(headers):
        width = f" width:{widths[index]};" if widths and index < len(widths) else ""
        header_cells.append(
            f'<th style="text-align:left; vertical-align:top; padding:8px 10px; border:1px solid #b8c2cc;'
            f' background:#e8eef5; color:#123b5d; font-size:13px;{width}">{_html_text(label)}</th>'
        )
    body_rows = []
    for row in rows:
        cells = "".join(
            f'<td style="text-align:left; vertical-align:top; padding:8px 10px; border:1px solid #cbd3da;'
            f' font-size:13px; line-height:1.4;">{_html_text(value)}</td>'
            for value in row
        )
        body_rows.append(f"<tr>{cells}</tr>")
    return (
        '<table role="presentation" style="border-collapse:collapse; width:100%; margin:8px 0 18px 0;'
        ' font-family:Arial,Helvetica,sans-serif;">'
        f"<thead><tr>{''.join(header_cells)}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"
    )


def _booking_section_html(
    bookings: list[AmenityBooking], start_date: date, end_date: date, error: str | None
) -> str:
    content = [
        '<h2 style="font-size:20px; line-height:1.25; color:#083b66; margin:0 0 6px 0;">Upcoming Amenity Bookings</h2>',
        f'<p style="margin:0 0 12px 0;">Period: {start_date:%Y-%m-%d} to {end_date:%Y-%m-%d}</p>',
    ]
    if error:
        content.append(
            f'<p style="color:#9a3412; font-weight:bold; margin:0;">Booking extraction failed: {_html_text(error)}</p>'
        )
    elif not bookings:
        content.append('<p style="margin:0;">No matching bookings found.</p>')
    else:
        rows = []
        for booking in bookings:
            time_range = f"{booking.start:%a %b %d, %H:%M} - "
            time_range += (
                f"{booking.end:%H:%M}"
                if booking.end.date() == booking.start.date()
                else f"{booking.end:%a %b %d, %H:%M}"
            )
            rows.append([
                booking.amenity,
                time_range,
                booking.requested_for,
                booking.description,
                booking.status,
            ])
        content.append(_html_table(
            ["Amenity", "Reservation Time", "Requested For", "Description", "Status"],
            rows,
            ["16%", "24%", "25%", "23%", "12%"],
        ))
    return "".join(content)


def _repair_section_html(cases: list[RepairCase], max_cases: int) -> str:
    content = [
        '<h2 style="font-size:20px; line-height:1.25; color:#083b66; margin:0 0 6px 0;">Repair Log</h2>',
        f'<p style="margin:0 0 16px 0;">Total open cases with valid reported dates: {len(cases)}</p>',
    ]
    if not cases:
        return "".join(content) + '<p style="margin:0;">No open repair cases found.</p>'

    category_order = list(dict.fromkeys(case.category for case in cases))
    for category in category_order:
        all_category_cases = sorted(
            (case for case in cases if case.category == category),
            key=lambda item: (item.reported_date, item.unit.casefold()),
        )
        category_cases = all_category_cases[:max_cases]
        if not category_cases:
            continue
        content.append(
            f'<h3 style="font-size:18px; line-height:1.3; color:#1f2937; margin:20px 0 6px 0;'
            f' font-weight:700;">{_html_text(category)}</h3>'
        )
        content.append(_html_table(
            ["Reported Date", "Unit", "Description", "Note"],
            [[case.reported_date.strftime("%Y-%m-%d"), case.unit, case.description, case.note]
             for case in category_cases],
            ["16%", "14%", "42%", "28%"],
        ))
        if len(all_category_cases) > max_cases:
            content.append(
                f'<p style="margin:-10px 0 18px 0; color:#4b5563;">'
                f'{len(all_category_cases) - max_cases} additional {_html_text(category)} case(s) are not shown.</p>'
            )
    return "".join(content)


def _bulk_section_html(requests: list[BulkItemRequest], enabled: bool) -> str:
    content = ['<h2 style="font-size:20px; line-height:1.25; color:#083b66; margin:0 0 8px 0;">Bulk Item Requests</h2>']
    if not enabled:
        return "".join(content) + '<p style="margin:0;">This section is not enabled yet.</p>'
    if not requests:
        return "".join(content) + '<p style="margin:0;">No open bulk item requests found.</p>'
    rows = [[
        request.reported_date.strftime("%Y-%m-%d") if request.reported_date else "Date not entered",
        request.property_name,
        request.unit,
        request.items,
        request.note,
    ] for request in requests]
    content.append(_html_table(
        ["Reported Date", "Property", "Unit", "Items", "Contact/Note"], rows,
        ["14%", "22%", "12%", "30%", "22%"],
    ))
    return "".join(content)


def build_email_html(
    cases: list[RepairCase],
    bookings: list[AmenityBooking],
    bulk_requests: list[BulkItemRequest],
    start_date: date,
    end_date: date,
    config: dict[str, Any],
    extraction_error: str | None = None,
) -> str:
    information_path = config.get("information_path") or config.get(
        "workbook", "Repair Log, Bulk Item Requests and Amenity Bookings.xlsx"
    )
    separator = '<div style="border-top:2px solid #d7dee5; margin:24px 0;"></div>'
    sections = [
        _booking_section_html(bookings, start_date, end_date, extraction_error),
        _repair_section_html(cases, int(config.get("max_cases", 20))),
        _bulk_section_html(bulk_requests, bool(config.get("bulk_items", {}).get("enabled", False))),
    ]
    return (
        '<!doctype html><html><body style="margin:0; padding:0; background:#f4f6f8;">'
        '<div style="max-width:960px; margin:0 auto; padding:24px; background:#ffffff;'
        ' color:#111827; font-family:Arial,Helvetica,sans-serif; font-size:14px; line-height:1.45;">'
        f'<h1 style="font-size:24px; line-height:1.25; color:#062f4f; margin:0 0 24px 0;">'
        f'Property Operations Daily Reminder - {date.today():%Y-%m-%d}</h1>'
        f'<p style="margin:-12px 0 24px 0; color:#374151;">'
        f'For more information, please go to {_file_link(str(information_path))}</p>'
        f'{separator.join(sections)}</div></body></html>'
    )


def load_config(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def application_folder() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def resolve_app_path(value: str, packaged_parent: bool = False) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    base = application_folder().parent if packaged_parent and getattr(sys, "frozen", False) else application_folder()
    return base / path


def send_email(config: dict[str, Any], subject: str, body: str, html_body: str) -> None:
    settings = config["email"]
    message = EmailMessage()
    message["From"] = settings["from"]
    message["To"] = ", ".join(settings["to"])
    if settings.get("cc"):
        message["Cc"] = ", ".join(settings["cc"])
    message["Subject"] = subject
    message.set_content(body)
    message.add_alternative(html_body, subtype="html")
    recipients = list(settings["to"]) + list(settings.get("cc", []))
    password = settings.get("password") or os.environ.get(settings.get("password_env", "SMTP_PASSWORD"))
    if settings.get("username") and not password:
        raise RuntimeError(f"Email password unavailable; set {settings.get('password_env', 'SMTP_PASSWORD')}.")
    with smtplib.SMTP(settings["smtp_server"], int(settings.get("smtp_port", 587)), timeout=30) as smtp:
        if settings.get("starttls", True):
            smtp.starttls()
        if settings.get("username"):
            smtp.login(settings["username"], password)
        smtp.send_message(message, to_addrs=recipients)


def main() -> int:
    parser = argparse.ArgumentParser(description="Refresh BuildingLink bookings and send the daily property operations reminder.")
    parser.add_argument("workbook", nargs="?", default=None)
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--preview", action="store_true")
    parser.add_argument("--output")
    parser.add_argument("--html-output")
    parser.add_argument("--skip-buildinglink", action="store_true")
    args = parser.parse_args()

    config = load_config(resolve_app_path(args.config))
    workbook_setting = args.workbook or config.get(
        "workbook", "Repair Log, Bulk Item Requests and Amenity Bookings.xlsx"
    )
    workbook_path = resolve_app_path(workbook_setting, packaged_parent=True)
    if not workbook_path.is_file():
        raise FileNotFoundError(
            f"Workbook not found: {workbook_path}. Check the 'workbook' value in config.json "
            "or pass the workbook path as the first command-line argument."
        )
    booking_config = config.get("buildinglink", {})
    start_date = date.today()
    end_date = start_date + timedelta(days=int(booking_config.get("days_ahead", 7)))
    bookings: list[AmenityBooking] = []
    extraction_error: str | None = None
    if not args.skip_buildinglink and booking_config.get("enabled", True):
        try:
            username = os.environ.get("BUILDINGLINK_USERNAME", "").strip()
            password = os.environ.get("BUILDINGLINK_PASSWORD", "")
            if not username or not password:
                raise RuntimeError("BUILDINGLINK_USERNAME or BUILDINGLINK_PASSWORD is not available.")
            bookings = extract_bookings(
                username, password, start_date, end_date,
                {str(item).casefold() for item in booking_config.get("amenities", [])},
                {str(item).casefold() for item in booking_config.get("statuses", [])},
                booking_config.get("property_name", "Example Condominium - DEMO 1001"),
                booking_config.get("url", "https://secure.buildinglink.com/v2/mgmt/amenities/listview.aspx"),
                headless=bool(booking_config.get("headless", True)),
                timeout=int(booking_config.get("login_timeout_seconds", 60)),
            )
            update_amenity_sheet(workbook_path, bookings, datetime.now())
        except Exception as exc:
            extraction_error = str(exc)

    cases = read_open_cases(workbook_path)
    bulk_requests = read_bulk_requests(workbook_path)
    body = build_email_body(cases, bookings, bulk_requests, start_date, end_date, config, extraction_error)
    html_body = build_email_html(cases, bookings, bulk_requests, start_date, end_date, config, extraction_error)
    if args.output:
        Path(args.output).write_text(body, encoding="utf-8")
    if args.html_output:
        Path(args.html_output).write_text(html_body, encoding="utf-8")
    if args.preview:
        print(body)
        return 0
    send_email(config, config.get("subject", "Property Operations Daily Reminder"), body, html_body)
    print(f"Sent reminder: {len(cases)} repair case(s), {len(bookings)} booking(s), {len(bulk_requests)} bulk request(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit("Public showcase: live worker execution is disabled.")
    raise SystemExit(main())
