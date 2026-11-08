"""JSON-lines process bridge. Existing matching logic lives in P2P_Registry_Sync.

stdout is exclusively structured events; a per-run file requests cancellation.
Nothing is saved before processing completes and the original is backed up.
"""
import argparse
import contextlib
import hashlib
import io
import json
import os
from pathlib import Path
import sys
import tempfile
from workbook_visibility import preserve_visibility, prepare_staged_file
import threading
import time
import uuid
from types import SimpleNamespace
from request_policy import Cooldown, MeasuredRequests
from datetime import datetime

PROTOCOL_OUT = sys.stdout
CANCEL = threading.Event()
CANCEL_FILE = None
RETRY_FILE = None
LOGIN_CREDENTIALS = None

if getattr(sys, "frozen", False):
    browser_dir = Path(sys.executable).parent / "browsers"
    if browser_dir.is_dir():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browser_dir)


class Cancelled(Exception):
    pass


def emit(kind, **values):
    PROTOCOL_OUT.write(json.dumps({"type": kind, **values}, ensure_ascii=False, default=str) + "\n")
    PROTOCOL_OUT.flush()


def check_cancel():
    if CANCEL.is_set() or (CANCEL_FILE and Path(CANCEL_FILE).exists()):
        raise Cancelled("Sync cancelled. The workbook was not changed.")


class LogWriter(io.TextIOBase):
    def __init__(self):
        self.pending = ""
        self.row_messages = []

    def write(self, value):
        self.pending += value
        while "\n" in self.pending:
            line, self.pending = self.pending.split("\n", 1)
            if line.strip():
                self.row_messages.append(line.strip())
                emit("log", message=line.strip())
        return len(value)

    def flush(self):
        if self.pending.strip():
            line, self.pending = self.pending.strip(), ""
            self.row_messages.append(line)
            emit("log", message=line)


def fingerprint(path):
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def backup_directory(path, configured=None):
    if configured and str(configured).strip():
        directory = Path(configured)
        if not directory.is_absolute():
            raise ValueError("The registry backup folder must be a full path.")
        if directory.exists() and not directory.is_dir():
            raise ValueError("The registry backup location is a file. Choose a folder instead.")
        return directory
    return Path(path).parent / "Registry Backups"


def registry_source_layout(sync, workbook):
    """Contractors and invoice numbers are input data, never sync outputs."""
    layout = {}
    for ws in workbook.worksheets:
        columns = sync.get_col_map(ws)
        protected = [columns[name] for name in (sync.H_VENDOR, sync.H_INVOICE) if name in columns]
        if protected:
            layout[ws.title] = (protected, [
                (row, col, ws.cell(row, col).value, ws.cell(row, col).data_type)
                for row in range(1, ws.max_row + 1) for col in protected
                if ws.cell(row, col).value is not None
            ])
    return layout


def assert_registry_source_unchanged(expected, actual):
    if expected != actual:
        raise RuntimeError("Contractor names or invoice numbers changed row positions during sync. Nothing was saved. The original workbook was preserved.")


def save_safely(sync, workbook, path, original_hash, backup_folder=None):
    """Stage a full zip before replacing; abort if another program changed input."""
    import shutil
    check_cancel()
    path = Path(path)
    if fingerprint(path) != original_hash:
        raise RuntimeError("The workbook changed during sync. Nothing was saved. Close Excel and run again.")
    from openpyxl import load_workbook
    original = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    try:
        source_layout = registry_source_layout(sync, original)
    finally:
        original.close()
    assert_registry_source_unchanged(source_layout, registry_source_layout(sync, workbook))
    backup_dir = backup_directory(path, backup_folder)
    backup = backup_dir / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}{path.suffix}"
    try:
        backup_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backup)
        if fingerprint(backup) != original_hash:
            raise RuntimeError("The workbook changed while the backup was being created. Nothing was saved.")
    except Exception:
        try:
            backup.unlink(missing_ok=True)
        except OSError:
            pass
        raise

    handle, staged = tempfile.mkstemp(prefix="registry-sync-", suffix=path.suffix, dir=path.parent)
    os.close(handle)
    try:
        prepare_staged_file(staged)
        workbook.save(staged)
        saved = load_workbook(staged, keep_vba=path.suffix.lower() == ".xlsm")
        try:
            assert_registry_source_unchanged(source_layout, registry_source_layout(sync, saved))
        finally:
            saved.close()
        if fingerprint(path) != original_hash:
            raise RuntimeError("The workbook changed before saving. Nothing was overwritten.")
        preserve_visibility(path, staged)
        os.replace(staged, path)
    finally:
        if os.path.exists(staged):
            os.unlink(staged)
    return str(backup)


def run(path, check_only=False, backup_folder=None):
    global LOGIN_CREDENTIALS
    started = time.monotonic()
    timings = {}
    requests = None
    from openpyxl import load_workbook
    import P2P_Registry_Sync as sync

    timing_reported = False

    def report_timing():
        nonlocal timing_reported
        if timing_reported:
            return
        timing_reported = True
        timings["total_seconds"] = round(time.monotonic() - started, 3)
        emit("log", message="Sync phase timing: " + json.dumps(timings, sort_keys=True))
        if requests is not None:
            emit("log", message="Sync API timing (excludes sign-in checks): " + json.dumps(requests.summary(), sort_keys=True))

    path = Path(path).resolve()
    if path.suffix.lower() not in (".xlsx", ".xlsm") or path.name.startswith("~$"):
        raise ValueError("Choose a valid .xlsx or .xlsm registry workbook.")
    if not check_only:
        backup_directory(path, backup_folder)
    if not path.is_file():
        raise FileNotFoundError("The selected workbook no longer exists. Browse to it again.")
    if path.with_name("~$" + path.name).exists() and not check_only:
        raise PermissionError("This workbook appears to be open in Excel. Close it before starting sync.")
    if not check_only:
        with open(path, "r+b"):
            pass
    emit("stage", stage="loading", message="Reading the selected registry workbook…")
    original_hash = fingerprint(path)
    workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    timings["workbook_seconds"] = round(time.monotonic() - started, 3)
    context = None
    log = LogWriter()
    try:
        sheet_rows = {}
        sheet_columns = {}
        for ws in workbook.worksheets:
            columns = sync.get_col_map(ws)
            sheet_columns[ws.title] = columns
            invoice_col = columns.get(sync.H_INVOICE)
            if invoice_col:
                sheet_rows[ws.title] = [row for row in range(2, ws.max_row + 1)
                    if not sync.is_empty(ws.cell(row, invoice_col).value)]
        total_rows = sum(map(len, sheet_rows.values()))
        if not total_rows:
            raise ValueError("No invoice rows were found. Check that the header row contains 'Invoice No'.")
        excluded_rows = sum(bool(sync.registry_skip_reason(workbook[name], row, sheet_columns[name]))
                            for name, rows in sheet_rows.items() for row in rows)
        emit("validated", total=total_rows, sheets=len(sheet_rows), excluded=excluded_rows,
             eligible=total_rows - excluded_rows, workbook=str(path), backup_folder=str(backup_directory(path, backup_folder)))
        if check_only:
            report_timing()
            emit("complete", check_only=True, reviewed=0, updated=0, attention=0,
                 total=total_rows, skipped=excluded_rows, saved=False, backup=None, workbook=str(path), stats={})
            return
        check_cancel()
        if excluded_rows == total_rows:
            report_timing()
            emit("complete", reviewed=total_rows, updated=0, attention=0, skipped=excluded_rows,
                 total=total_rows, saved=False, backup=None, workbook=str(path),
                 stats={**sync.new_sync_stats(), "excluded_rows": excluded_rows})
            return
        # All desktop app instances use the same lock and the existing browser profile.
        os.makedirs(sync.APP_DATA_DIR, exist_ok=True)
        lock_path = Path(sync.APP_DATA_DIR) / "desktop-sync.lock"
        with open(lock_path, "a+b") as lock:
            import msvcrt
            lock.seek(0)
            if not lock.read(1):
                lock.write(b"0")
                lock.flush()
            lock.seek(0)
            try:
                msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                raise RuntimeError("Another app sync is running. Finish or cancel it before starting again.")
            try:
                with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
                    from playwright.sync_api import sync_playwright
                    os.makedirs(sync.SESSION_DIR, exist_ok=True)
                    with sync_playwright() as playwright:
                        import session_transport
                        auth_started = time.monotonic()
                        with session_transport.open_session(playwright, sync, LOGIN_CREDENTIALS, RETRY_FILE, emit, check_cancel) as (page, headers):
                            timings["sign_in_seconds"] = round(time.monotonic() - auth_started, 3)
                            requests = MeasuredRequests(page.request, sync, Cooldown.for_sync(sync), check_cancel)
                            page = SimpleNamespace(request=requests)
                            index_started = time.monotonic()
                            LOGIN_CREDENTIALS = None
                            check_cancel()
                            emit("stage", stage="lookup", message="Connected to Yardi. Loading invoice and PO indexes…")
                            irgrid = sync.preload_irgrid(page, headers)
                            check_cancel()
                            pogrid = sync.preload_pogrid(page, headers)
                            check_cancel()
                            timings["index_seconds"] = round(time.monotonic() - index_started, 3)
                            emit("log", message="Automatic Account Code checking was rolled back. Existing codes, including split-invoice rows, are kept; no coding-only detail requests are made.")
                            rows_started = time.monotonic()
                            totals = sync.new_sync_stats()
                            reviewed = updated = attention = 0
                            caches = ({}, {}, {}, {})
                            for ws in workbook.worksheets:
                                if ws.title not in sheet_rows:
                                    emit("log", message=f"Skipped sheet '{ws.title}': no Invoice No header.")
                                    continue
                                columns = sync.get_col_map(ws)
                                if sync.H_STATUS not in columns:
                                    emit("log", message=f"'{ws.title}' has no P2P Status column; other supported fields can still be filled.")
                                tracked = [columns[name] for name in
                                    (sync.H_STATUS, sync.H_PO, sync.H_IR, sync.H_CH_NUM, sync.H_CH_DATE, sync.H_PAYMENT_METHOD, sync.H_ACCOUNT)
                                    if name in columns]
                                pending = None
                                before = ()
                                pending_skip_reason = ""
                                issues = {}

                                def report_attention(row, reason):
                                    issues.setdefault(row, []).append(reason)

                                def finish_row():
                                    nonlocal pending, reviewed, updated, attention
                                    requests.check()
                                    if pending is None:
                                        return
                                    log.flush()
                                    after = tuple(ws.cell(pending, col).value for col in tracked)
                                    changed = before != after
                                    reason = " ".join(dict.fromkeys(issues.get(pending, [])))
                                    needs_review = bool(reason)
                                    status_col = columns.get(sync.H_STATUS)
                                    status = str(ws.cell(pending, status_col).value or "") if status_col else ""
                                    reviewed += 1
                                    updated += int(changed)
                                    attention += int(needs_review)
                                    def value(header):
                                        col = columns.get(header)
                                        return str(ws.cell(pending, col).value or "") if col else ""
                                    emit("row", sheet=ws.title, row=pending,
                                         vendor=value(sync.H_VENDOR), invoice=value(sync.H_INVOICE),
                                         ir=value(sync.H_IR), po=value(sync.H_PO), status=status,
                                         result="Skipped" if pending_skip_reason else "Needs review" if needs_review else "Updated" if changed else "Unchanged",
                                         skipped=bool(pending_skip_reason), skip_reason=pending_skip_reason,
                                         reason=reason, changed=changed, reviewed=reviewed, updated=updated,
                                         attention=attention, total=total_rows)
                                    pending = None

                                def start_row(row):
                                    nonlocal pending, before, pending_skip_reason
                                    finish_row()
                                    check_cancel()
                                    pending = row
                                    pending_skip_reason = sync.registry_skip_reason(ws, row, columns)
                                    before = tuple(ws.cell(row, col).value for col in tracked)
                                    log.row_messages.clear()
                                    emit("stage", stage="syncing", message=f"Reviewing {ws.title} · row {row}")

                                stats = sync.sync_sheet(ws, ws.title, page, headers,
                                    irgrid, pogrid, *caches, progress_callback=start_row, attention_callback=report_attention)
                                finish_row()
                                for key in totals:
                                    totals[key] += stats.get(key, 0)
                            timings["rows_seconds"] = round(time.monotonic() - rows_started, 3)
                            requests.check()
                    check_cancel()
                    save_started = time.monotonic()
                    backup = None
                    if updated:
                        emit("stage", stage="saving", message="Creating a backup and saving the registry. Please keep Excel closed.")
                        backup = save_safely(sync, workbook, path, original_hash, backup_folder)
                    timings["save_seconds"] = round(time.monotonic() - save_started, 3)
                    report_timing()
                    emit("complete", performance=requests.summary(), reviewed=reviewed, updated=updated, attention=attention, skipped=totals.get("excluded_rows", 0),
                         total=total_rows, saved=bool(updated), backup=backup, workbook=str(path), stats=totals)
            finally:
                lock.seek(0)
                msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        workbook.close()
        report_timing()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook")
    parser.add_argument("--backup-folder", help="Optional full path for original registry backups")
    parser.add_argument("--check-browser", action="store_true")
    parser.add_argument("--check-only", action="store_true")
    parser.add_argument("--cancel-file")
    parser.add_argument("--retry-file")
    parser.add_argument("--login-stdin", action="store_true")
    args = parser.parse_args()
    global CANCEL_FILE, RETRY_FILE, LOGIN_CREDENTIALS
    CANCEL_FILE = args.cancel_file
    RETRY_FILE = args.retry_file
    writer = LogWriter()
    try:
        if args.login_stdin:
            try:
                payload = json.loads(sys.stdin.readline(32768))
                if not isinstance(payload, dict):
                    raise ValueError()
                if payload:
                    if not all(isinstance(payload.get(key), str) and payload[key] for key in ("origin", "username", "password")):
                        raise ValueError()
                    LOGIN_CREDENTIALS = payload
            except Exception:
                raise ValueError("The sign-in details could not be received. Re-enter them in the app.") from None
        with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
            if args.check_browser:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch(channel="chromium", headless=True)
                    try:
                        page = browser.new_page()
                        page.set_content("<title>Bundled browser check</title>")
                        if page.title() != "Bundled browser check":
                            raise RuntimeError("Bundled browser check failed.")
                        emit("browser_ready", message="Bundled Chromium launched successfully.")
                    finally:
                        browser.close()
            elif args.workbook:
                run(args.workbook, args.check_only, args.backup_folder)
            else:
                raise ValueError("A workbook must be selected.")
        return 0
    except Cancelled as exc:
        emit("cancelled", message=str(exc))
        return 2
    except (Exception, SystemExit) as exc:
        writer.flush()
        message = str(exc)
        if isinstance(exc, ModuleNotFoundError):
            message = "A required backend component is missing. Re-extract the complete application folder. " + message
        if isinstance(exc, PermissionError):
            message = "Cannot access the workbook or backup folder. Close Excel and check write permissions. " + message
        emit("error", message=message or "The sync backend could not start. Check the private Yardi configuration.")
        return 1


if __name__ == "__main__":
    raise SystemExit("Public showcase: live worker execution is disabled.")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())

