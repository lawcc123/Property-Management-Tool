"""Reservation/operations bridge. No automatic mail from preview; secrets only on stdin."""
import argparse
import contextlib
from dataclasses import asdict
from datetime import datetime, timedelta
from email.message import EmailMessage
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import re
import shutil
import smtplib
import sys
import tempfile
from workbook_visibility import preserve_visibility, prepare_staged_file
from zoneinfo import ZoneInfo

import operations_legacy as legacy
from buildinglink_extractor import extract_bookings, SHARED_PROPERTY


class Cancelled(Exception):
    pass


def emit(kind, **values):
    print(json.dumps(dict(type=kind, **values), default=str), flush=True)


def digest(path):
    with open(path, 'rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def fingerprint(settings):
    return hashlib.sha256(json.dumps(settings, sort_keys=True).encode()).hexdigest()


def safe_error(error, secrets=()):
    text = str(error)
    for secret in secrets:
        if secret:
            text = text.replace(secret, '[redacted]')
    text = re.split(r'Raw rows:|page text:|Visible inputs:|Call log:', text, flags=re.I)[0]
    return type(error).__name__ + ': ' + text[:650]


def validate(settings, sending=False):
    workbook = Path(settings.get('Workbook', ''))
    if not workbook.is_absolute() or workbook.suffix.lower() != '.xlsx':
        raise ValueError('Choose a full path to an .xlsx operations workbook.')
    if not workbook.is_file():
        raise ValueError('The operations workbook is unavailable. Check the network share.')
    if not 0 <= int(settings.get('DaysAhead', 7)) <= 90:
        raise ValueError('Days ahead must be between 0 and 90.')
    if sending:
        if not settings.get('Recipients') or not settings.get('From') or not settings.get('Subject'):
            raise ValueError('Set the sender, recipients and subject before sending.')
        for address in [settings['From']] + settings['Recipients'] + settings.get('Cc', []):
            if not re.fullmatch(r'[^\s<>@,;]+@[^\s<>@,;]+\.[^\s<>@,;]+', address):
                raise ValueError('Use complete email addresses without display names.')
        if any('\r' in str(settings.get(k, '')) or '\n' in str(settings.get(k, '')) for k in ['From', 'Subject', 'SmtpServer']):
            raise ValueError('Email fields must not contain line breaks.')
        if not settings.get('SmtpServer') or not 1 <= int(settings.get('SmtpPort', 587)) <= 65535:
            raise ValueError('Set the SMTP server and port.')
    return workbook


def retrieve(settings, credentials, progress):
    username, password = credentials.get('username', ''), credentials.get('password', '')
    if not username or not password:
        raise ValueError('Save a BuildingLink login in Manage credentials first.')
    start = datetime.now(ZoneInfo('America/Example City')).date()
    for attempt in range(2):
        try:
            return extract_bookings(username, password, start, start + timedelta(days=settings.get('DaysAhead', 7)),
                set(), {'requested', 'approved'}, SHARED_PROPERTY,
                headless=True, timeout=60, progress=progress)
        except Cancelled:
            raise
        except Exception as error:
            if attempt or 'Connection closed while reading from the driver' not in str(error):
                raise
            progress('Browser driver disconnected; retrying once with a fresh session')
    raise RuntimeError('Booking extraction did not finish.')


def sheet_checks(local, settings):
    book = legacy.load_workbook(local, read_only=True, data_only=True)
    warnings = []
    try:
        bulk_sheets = [s for s in book if 'bulk' in s.title.casefold()]
        if not bulk_sheets:
            warnings.append('No bulk-item worksheet found; this is not a confirmed empty request list.')
        for sheet in bulk_sheets:
            if not {'unit', 'items'}.issubset(legacy._headers(sheet)):
                warnings.append('Bulk worksheet "' + sheet.title + '" needs Unit and Items headers in its first row.')
        if not any({'unit', 'description'}.issubset(legacy._headers(s)) for s in book):
            warnings.append('No repair worksheet with Unit and Description headers was found.')
    finally:
        book.close()
    return warnings


@contextlib.contextmanager
def locked_source(source):
    """Deny concurrent readers/writers while allowing atomic replacement on Windows."""
    if os.name != 'nt':
        with open(source, 'rb') as stream:
            yield stream
        return
    import ctypes
    from ctypes import wintypes
    import msvcrt
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    handle = create(str(source), 0x80000000, 0x4, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY)
    except Exception:
        close = kernel.CloseHandle
        close.argtypes = [wintypes.HANDLE]
        close(handle)
        raise
    with os.fdopen(fd, 'rb') as stream:
        yield stream


def publish(local, source, expected_hash, progress):
    if digest(source) != expected_hash:
        raise RuntimeError('The workbook changed while bookings were retrieved. Refresh again; no workbook was overwritten.')
    progress('Saving the Amenity Booking worksheet')
    # One rotating recovery copy, not a new timestamped workbook on every run.
    backup_dir = source.parent / 'Operations reminder backups'
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / (source.stem + '-before-booking-refresh.xlsx')
    fd, stage_name = tempfile.mkstemp(prefix='operations-', suffix='.xlsx', dir=source.parent)
    os.close(fd)
    stage = Path(stage_name)
    backup_stage = backup.with_suffix('.tmp')
    try:
        prepare_staged_file(stage)
        shutil.copyfile(local, stage)
        with locked_source(source) as original:
            if hashlib.file_digest(original, 'sha256').hexdigest() != expected_hash:
                raise RuntimeError('The workbook changed before save. No workbook was overwritten.')
            original.seek(0)
            with open(backup_stage, 'wb') as target:
                shutil.copyfileobj(original, target)
            os.replace(backup_stage, backup)
        # Windows replacement needs its own access to the destination. Release the
        # sharing guard, recheck the source, then atomically publish. An open Excel
        # workbook also prevents replacement through its Windows sharing rules.
        if digest(source) != expected_hash:
            raise RuntimeError('The workbook changed before publication. No workbook was overwritten.')
        preserve_visibility(source, stage)
        os.replace(stage, source)
    finally:
        stage.unlink(missing_ok=True)
        backup_stage.unlink(missing_ok=True)


def preview(settings, credentials, progress):
    source = validate(settings)
    progress('Checking operations workbook', 3)
    original_hash = digest(source)
    start = datetime.now(ZoneInfo('America/Example City')).date()
    end = start + timedelta(days=settings.get('DaysAhead', 7))
    with tempfile.TemporaryDirectory(prefix='PropertyManager-operations-') as temporary:
        local = Path(temporary) / 'operations.xlsx'
        shutil.copyfile(source, local)
        if digest(local) != original_hash:
            raise RuntimeError('Workbook changed while it was being read. Refresh again.')
        warnings = sheet_checks(local, settings)
        progress('Reading open repair cases', 12)
        repairs = legacy.read_open_cases(local)
        progress('Reading open bulk-item requests', 22)
        bulk = legacy.read_bulk_requests(local)
        bookings, booking_error, save_error = [], None, None
        try:
            bookings = retrieve(settings, credentials, lambda message: progress(message, None))
        except Cancelled:
            raise
        except Exception as error:
            booking_error = safe_error(error, (credentials.get('username'), credentials.get('password')))
            warnings.append('Booking refresh failed: ' + booking_error + '. Previous Excel bookings were preserved.')
        if booking_error is None:
            try:
                progress('Preparing updated booking worksheet', 78)
                legacy.update_amenity_sheet(local, bookings, datetime.now(ZoneInfo('America/Example City')).replace(tzinfo=None))
                publish(local, source, original_hash, progress)
            except Cancelled:
                raise
            except Exception as error:
                save_error = safe_error(error)
                warnings.append('Bookings retrieved, but Excel was not updated: ' + save_error)
        progress('Preparing reminder preview', 92)
        config = {'max_cases': 20, 'workbook': str(source), 'bulk_items': {'enabled': True}}
        body = legacy.build_email_body(repairs, bookings, bulk, start, end, config, booking_error)
        html = legacy.build_email_html(repairs, bookings, bulk, start, end, config, booking_error)
        if warnings:
            import html as html_module
            body = 'RUN NOTICES\n' + '\n'.join(warnings) + '\n\n' + body
            html = '<div style="padding:12px;background:#fff4df">' + '<br>'.join(html_module.escape(w) for w in warnings) + '</div>' + html
        progress('Preview ready', 100)
        return {'id': os.urandom(16).hex(), 'created': datetime.now(ZoneInfo('America/Example City')).isoformat(),
            'settings_fingerprint': fingerprint(settings), 'workbook_hash': digest(source),
            'start': str(start), 'end': str(end), 'repairs': [asdict(r) for r in repairs], 'bulk': [asdict(r) for r in bulk],
            'bookings': [asdict(b) for b in bookings], 'warnings': warnings, 'booking_error': booking_error,
            'save_error': save_error, 'body': body, 'html': html}


def send(settings, prepared, credentials, progress, transport=None):
    workbook = validate(settings, sending=True)
    if fingerprint(settings) != prepared.get('settings_fingerprint'):
        raise ValueError('Reminder settings changed. Refresh the preview before sending.')
    created = datetime.fromisoformat(prepared['created'])
    age = datetime.now(ZoneInfo('America/Example City')) - created
    if age.total_seconds() < -60 or age > timedelta(minutes=30):
        raise ValueError('The preview is over 30 minutes old. Refresh it before sending.')
    if digest(workbook) != prepared['workbook_hash']:
        raise ValueError('The workbook changed after preview. Refresh before sending.')
    if prepared.get('warnings') and not settings.get('AllowPartial', False):
        raise ValueError('The preview has warnings. Resolve them or explicitly enable sending reminders with warnings.')
    username, password = credentials.get('username', ''), credentials.get('password', '')
    if not username or not password:
        raise ValueError('Save the reminder email account/app password in Manage credentials.')
    message = EmailMessage()
    message['From'], message['Subject'] = settings['From'], settings['Subject']
    message['To'] = ', '.join(settings['Recipients'])
    if settings.get('Cc'):
        message['Cc'] = ', '.join(settings['Cc'])
    message.set_content(prepared['body'])
    message.add_alternative(prepared['html'], subtype='html')
    progress('Connecting to the mail server', 25)
    with (transport or smtplib.SMTP)(settings['SmtpServer'], int(settings.get('SmtpPort', 587)), timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(username, password)
        progress('Submitting email; do not retry if the outcome becomes uncertain', 70, submitting=True)
        rejected = smtp.send_message(message, to_addrs=settings['Recipients'] + settings.get('Cc', []))
        if rejected:
            raise RuntimeError('The mail server rejected some recipients. Other recipients may have received it. Do not resend the whole list automatically.')
    progress('Mail server accepted the reminder', 100)
    return {'accepted': True, 'message': 'Accepted by the mail server; delivery is not confirmed.'}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--action', choices=['preview', 'send', 'campaign-delivery'], required=True)
    parser.add_argument('--cancel', required=True)
    args = parser.parse_args()
    packet = json.loads(sys.stdin.readline())
    if args.action == 'campaign-delivery':
        from campaign_mail import run as campaign_run
        return campaign_run(packet, emit)
    secrets = [v for c in packet.get('credentials', {}).values() if isinstance(c, dict) for v in c.values() if isinstance(v, str)]
    stage = 'Starting'
    def progress(message, percent=None, **extra):
        nonlocal stage
        if Path(args.cancel).exists():
            raise Cancelled('Cancelled before the next operation; previous completed workbook updates remain saved.')
        stage = message
        emit('status', message=message, percent=percent, **extra)
    try:
        if args.action == 'preview':
            result = preview(packet['settings'], packet.get('credentials', {}).get('buildinglink', {}), progress)
        else:
            result = send(packet['settings'], packet['preview'], packet.get('credentials', {}).get('email', {}), progress)
        emit('complete', success=True, result=result)
        return 0
    except Cancelled as error:
        emit('complete', success=False, cancelled=True, message=str(error), stage=stage)
        return 2
    except Exception as error:
        emit('complete', success=False, message=safe_error(error, secrets), stage=stage, diagnostics={'python':sys.version.split()[0],'playwright':importlib.metadata.version('playwright')})
        return 1


if __name__ == '__main__':
    raise SystemExit("Public showcase: live worker execution is disabled.")
    raise SystemExit(main())
