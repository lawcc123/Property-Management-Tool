"""
P2P_Registry_Sync.py

Syncs Yardi P2P milestone data into the selected registry workbook.

Columns filled / updated:
  Account Code is preserved; automatic coding checks have been withdrawn.
  P2P Status  <- ir/milestones or po/milestones  (only if column exists)
  P.O. #      <- IRGrid API                      (only fills empty cells)
  P2P IR #    <- IRGrid API                      (only fills empty cells)
  Ch.#        <- ir/milestones Paid milestone     (only fills empty cells)
  Ch. Date    <- ir/milestones Paid milestone     (only fills empty cells)

All column positions are detected dynamically from the header row so the
script keeps working even if columns are reordered later.

What it NEVER does:
  - Delete or move any row (empty rows kept for manager's visual layout)
  - Insert the P2P Status column if it is missing
  - Overwrite PO, IR, Ch.#, Ch. Date cells that already have a value
  - Change formatting, column widths, or styles on existing cells
  - Save changes without first backing up the selected input workbook

Run:  python P2P_Registry_Sync.py
"""

import json
import html
import sys
import os
import re
import shutil
import time
import base64
import unicodedata
from difflib import SequenceMatcher
from datetime import datetime
from urllib.parse import parse_qsl, urlsplit

from request_policy import RequestStopped
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR      = get_base_dir()
CONFIG_PATH   = os.path.join(BASE_DIR, "p2p_private_config.example.json")
LOCAL_APP_DATA = os.environ.get(
    "LOCALAPPDATA",
    os.path.join(os.path.expanduser("~"), "AppData", "Local"),
)
APP_DATA_DIR  = os.path.join(LOCAL_APP_DATA, "PropertyManagerDemo_P2P")
SESSION_DIR   = os.path.join(APP_DATA_DIR, "browser_session")
AUTH_STATE_PATH = os.path.join(APP_DATA_DIR, "auth_state.json")
DEFAULT_REGISTRY_PATH = os.path.join(BASE_DIR, "01. Cheque Registry_Year 2026.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "Output Excel File")
PAYMENT_DIAGNOSTIC_FLAG = "--diagnose-payment-method"
HISTORICAL_EFT_REFRESH_FLAG = "--refresh-historical-eft"
COMMAND_FLAGS = (PAYMENT_DIAGNOSTIC_FLAG, HISTORICAL_EFT_REFRESH_FLAG)
PAYMENT_DIAGNOSTIC_FIELD_HINTS = (
    "payment", "method", "check", "cheque", "eft", "electronic", "pmt", "ach",
)


def load_private_config():
    if not os.path.isfile(CONFIG_PATH):
        raise FileNotFoundError(
            f"Missing private config file:\n  {CONFIG_PATH}\n\n"
            "Copy p2p_private_config.example.json, fill in your private values, "
            "and keep p2p_private_config.json out of public repos."
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
        return json.load(config_file)


def wait_after_startup_error():
    """Keep a double-clicked Windows executable open so the error is visible."""
    if not getattr(sys, "frozen", False):
        return
    try:
        input("\nPress Enter to close this window...")
    except (EOFError, KeyboardInterrupt):
        pass


try:
    PRIVATE_CONFIG = load_private_config()
except (OSError, json.JSONDecodeError) as exc:
    print("\nP2P Registry Sync could not start.")
    print("-" * 44)
    print(exc)
    wait_after_startup_error()
    raise SystemExit(1)

if not os.path.isfile(DEFAULT_REGISTRY_PATH):
    for filename in os.listdir(BASE_DIR):
        if filename.startswith("~$"):
            continue
        if filename.lower().endswith((".xlsx", ".xlsm")):
            DEFAULT_REGISTRY_PATH = os.path.join(BASE_DIR, filename)
            break

LOGIN_URL        = PRIVATE_CONFIG["login_url"]
API_BASE         = PRIVATE_CONFIG["api_base"].rstrip("/")
P2P_DATABASE     = PRIVATE_CONFIG["database"]
P2P_ORIGIN       = PRIVATE_CONFIG["origin"]
P2P_REFERER      = PRIVATE_CONFIG["referer"]
IRGRID_URL       = f"{API_BASE}/IRTileData/IRGrid"
IR_DETAIL_URL    = f"{API_BASE}/ir"
IR_MILESTONE_URL = f"{API_BASE}/ir/milestones"
PO_MILESTONE_URL = f"{API_BASE}/po/milestones"
POGRID_URL       = f"{API_BASE}/po/charts/data"
BULK_IRGRID_RECORD_LIMIT = 1000
BULK_PO_RECORD_LIMIT = 1000

# Header names in the registry — used to find columns dynamically
H_CH_DATE   = "Ch. Date (DD/MM/YYYY)"
H_CH_NUM    = "Ch.#"
H_PAYMENT_METHOD = "Payment Method"
H_ACCOUNT = "Account Code"
H_STATUS    = "P2P Status"
H_PO        = "P.O. #"
H_IR        = "P2P IR #"
H_INVOICE   = "Invoice No"
H_VENDOR    = "Contractor Name"
H_AMOUNT    = "Invoice Amount"
H_INVOICE_DATE = "Invoice Date (DD/MM/YYYY)"
H_PROPERTY  = "Property"
H_REMARKS   = "Remarks"

IRGRID_INVOICE_FIELDS = ("Invoice", "InvoiceNumber", "InvoiceNo", "InvoiceNum", "InvNumber")
IRGRID_IR_FIELDS = ("Inv_hmy", "IR_hmy", "IR", "InvoiceNumber")
IRGRID_PO_FIELDS = ("PO_hmy", "PO", "PONumber", "POId")
IRGRID_STATUS_FIELDS = ("Status", "workflowStepName", "WorkflowStatus")
IRGRID_VENDOR_FIELDS = (
    "PayeeName", "payeeName", "VendorName", "vendorName", "Payee", "Vendor",
    "ContractorName", "Contractor",
)
IRGRID_AMOUNT_FIELDS = (
    "Amount", "InvoiceAmount", "TotalAmount", "GrossAmount", "BaseAmount",
)
IRGRID_DATE_FIELDS = ("InvoiceDate", "InvDate", "Date")
IRGRID_PROPERTY_FIELDS = (
    "PropertyCode", "propertyCode", "Property", "PropCode", "PropertyName",
)

POGRID_PO_FIELDS = (
    "PO", "POCode", "PONumber", "PONum", "PONo", "PO_hmy", "POId",
)
POGRID_DESCRIPTION_FIELDS = (
    "PODesc", "Description", "PODescription", "PoDescription", "Desc", "Notes",
)
MILESTONE_DATE_FIELDS = (
    "RevisedDate", "DtCompleted", "Date", "CompletedDate", "DateCompleted",
    "DtComplete", "DtCompletion", "CompletionDate", "ActualDate",
)


def clean_dropped_path(raw_path):
    """Normalize a path pasted by terminal drag-and-drop."""
    path = str(raw_path or "").strip()
    if path.startswith("& "):
        path = path[2:].strip()
    return path.strip().strip('"').strip("'")


def get_registry_path():
    """
    Accept an Excel path from argv or terminal drag-and-drop.
    Pressing Enter keeps the original default workbook path.
    """
    path_args = [arg for arg in sys.argv[1:] if arg not in COMMAND_FLAGS]
    if path_args:
        return clean_dropped_path(" ".join(path_args))

    print("Drag and drop the Excel registry file into this terminal, then press Enter.")
    print(f"Press Enter without a file to use: {os.path.basename(DEFAULT_REGISTRY_PATH)}")
    selected = clean_dropped_path(input("Excel file: "))
    return selected or DEFAULT_REGISTRY_PATH


def is_payment_method_diagnostic():
    return PAYMENT_DIAGNOSTIC_FLAG in sys.argv[1:]


def is_historical_eft_refresh():
    return HISTORICAL_EFT_REFRESH_FLAG in sys.argv[1:]


def get_backup_path(registry_path):
    """Create a timestamped backup path for the original input workbook."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    _, ext = os.path.splitext(os.path.basename(registry_path))
    return os.path.join(OUTPUT_DIR, f"back up_{ts}{ext}")


def backup_existing_workbook(registry_path):
    backup_path = get_backup_path(registry_path)
    shutil.copy2(registry_path, backup_path)
    return backup_path

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def restore_auth_state(context):
    """Restore saved cookies, including Yardi's non-persistent JWT cookie."""
    if not os.path.isfile(AUTH_STATE_PATH):
        return False
    try:
        with open(AUTH_STATE_PATH, "r", encoding="utf-8") as state_file:
            state = json.load(state_file)

        allowed_cookie_fields = {
            "name", "value", "url", "domain", "path", "expires",
            "httpOnly", "secure", "sameSite", "partitionKey",
        }
        cookies = [
            {
                key: value
                for key, value in cookie.items()
                if key in allowed_cookie_fields
            }
            for cookie in state.get("cookies", [])
        ]
        if not cookies:
            return False
        context.add_cookies(cookies)
        return True
    except (OSError, json.JSONDecodeError, TypeError, ValueError) as exc:
        print(f"Could not restore saved login state: {exc}")
        return False


def save_auth_state(context):
    """Save cookies and local storage for reuse on the next run."""
    try:
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        context.storage_state(path=AUTH_STATE_PATH)
        return True
    except OSError as exc:
        print(f"Could not save login state: {exc}")
        return False


def get_auth_headers(context):
    all_cookies = {c["name"]: c["value"] for c in context.cookies()}
    jwt_token = all_cookies.get(".JWTAUTH")
    if not jwt_token:
        return None
    try:
        payload_b64 = jwt_token.split(".")[1]
        payload_b64 += "=" * (4 - len(payload_b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_b64))
        if payload.get("exp", 0) < time.time():
            return None
        login_guid = payload.get("UserLoginGUID", "")
        role = payload.get(
            "http://schemas.microsoft.com/ws/2008/06/identity/claims/role", ""
        )
    except RequestStopped:
        raise
    except Exception:
        return None
    return {
        "Authorization":   f"Bearer {jwt_token}",
        "Database":        P2P_DATABASE,
        "Loginguid":       login_guid,
        "Role":            role,
        "Content-Type":    "application/json;charset=UTF-8",
        "Accept":          "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Origin":          P2P_ORIGIN,
        "Referer":         P2P_REFERER,
        "Sec-Fetch-Dest":  "empty",
        "Sec-Fetch-Mode":  "cors",
        "Sec-Fetch-Site":  "same-origin",
    }

def wait_for_jwt(context, timeout_seconds=300):
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        headers = get_auth_headers(context)
        if headers:
            return headers
        time.sleep(1)
    return None

def ensure_authenticated(page, context):
    headers = get_auth_headers(context)
    if headers:
        print("Valid session found — skipping login.\n")
        return headers
    print("\n  Browser opened — please log in to continue.")
    print("  After logging in, click 'Elevate (Production)' to open P2P.")
    print("  The script will continue automatically once you are in.\n")
    page.goto(LOGIN_URL)
    headers = wait_for_jwt(context, timeout_seconds=300)
    if headers:
        print("  Login detected — continuing.\n")
        return headers
    raise TimeoutError("Login timed out after 5 minutes.")


# ---------------------------------------------------------------------------
# Payment-method API discovery
# ---------------------------------------------------------------------------

def _diagnostic_endpoint(url):
    """Return a same-tenant endpoint without query values or credentials."""
    parsed = urlsplit(str(url or ""))
    allowed_host = urlsplit(P2P_ORIGIN).netloc.lower()
    if not parsed.netloc or parsed.netloc.lower() != allowed_host:
        return None
    query_keys = sorted({key for key, _ in parse_qsl(parsed.query)})
    result = parsed.path or "/"
    if query_keys:
        result += "?" + "&".join(f"{key}=<redacted>" for key in query_keys)
    return result


def _diagnostic_top_level_keys(payload):
    if isinstance(payload, dict):
        return sorted(str(key) for key in payload.keys())
    if isinstance(payload, list):
        object_item = next((item for item in payload if isinstance(item, dict)), None)
        if object_item is not None:
            return sorted(str(key) for key in object_item.keys())
        return [f"<list:{len(payload)}>"]
    return [f"<{type(payload).__name__}>"]


def _safe_diagnostic_value(value):
    if value is None or isinstance(value, (bool, int, float)):
        return value
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return text[:160] + ("..." if len(text) > 160 else "")


def _looks_like_payment_value(value):
    if not isinstance(value, str):
        return False
    normalized = re.sub(r"\s+", " ", value.strip().lower())
    return bool(re.search(
        r"\b(eft|ach|cheque|check|electronic funds? transfer|wire transfer)\b",
        normalized,
    ))


def _find_payment_fields(payload, path="$", results=None, depth=0):
    """Find payment-related JSON fields without recording unrelated values."""
    if results is None:
        results = []
    if depth > 12 or len(results) >= 100:
        return results

    if isinstance(payload, dict):
        for key, value in payload.items():
            key_text = str(key)
            child_path = f"{path}.{key_text}"
            normalized_key = re.sub(r"[^a-z0-9]", "", key_text.lower())
            key_matches = any(
                hint in normalized_key
                for hint in PAYMENT_DIAGNOSTIC_FIELD_HINTS
            )
            if key_matches or _looks_like_payment_value(value):
                if isinstance(value, (dict, list)):
                    results.append({
                        "path": child_path,
                        "value": f"<{type(value).__name__}>",
                    })
                else:
                    results.append({
                        "path": child_path,
                        "value": _safe_diagnostic_value(value),
                    })
            _find_payment_fields(value, child_path, results, depth + 1)
    elif isinstance(payload, list):
        for index, item in enumerate(payload[:100]):
            _find_payment_fields(item, f"{path}[{index}]", results, depth + 1)
            if len(results) >= 100:
                break
    return results


class PaymentMethodDiagnostic:
    """Collect a sanitized summary of same-tenant XHR/fetch JSON responses."""

    def __init__(self):
        self.records = []
        self._seen = set()

    def capture(self, response):
        try:
            if response.request.resource_type not in ("xhr", "fetch"):
                return
            endpoint = _diagnostic_endpoint(response.url)
            if endpoint is None:
                return
            content_type = response.headers.get("content-type", "").lower()
            if "json" not in content_type:
                return
            payload = response.json()
            fields = _find_payment_fields(payload)
            record = {
                "method": response.request.method,
                "endpoint": endpoint,
                "status": response.status,
                "topLevelKeys": _diagnostic_top_level_keys(payload),
                "paymentFields": fields,
            }
            signature = json.dumps(record, sort_keys=True, default=str)
            if signature in self._seen or len(self.records) >= 200:
                return
            self._seen.add(signature)
            self.records.append(record)
            if fields:
                print(f"  Candidate response: {record['method']} {endpoint}")
                for field in fields[:10]:
                    print(f"    {field['path']} = {field['value']}")
        except RequestStopped:
            raise
        except Exception as exc:
            # A failed or non-JSON response is not relevant to discovery.
            print(f"  Diagnostic skipped one response: {exc}")

    def save(self):
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(
            APP_DATA_DIR, f"payment_method_diagnostic_{timestamp}.json"
        )
        report = {
            "createdAt": datetime.now().isoformat(timespec="seconds"),
            "privacy": (
                "Query values, cookies, authorization headers, request bodies, "
                "and unrelated response values are not recorded."
            ),
            "responses": self.records,
        }
        with open(output_path, "w", encoding="utf-8") as output_file:
            json.dump(report, output_file, indent=2, ensure_ascii=False)
        return output_path


def run_payment_method_diagnostic(playwright):
    """Observe API responses while the user views EFT and cheque invoices."""
    os.makedirs(SESSION_DIR, exist_ok=True)
    try:
        context = playwright.chromium.launch_persistent_context(
            SESSION_DIR,
            channel="chrome",
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_https_errors=True,
        )
    except RequestStopped:
        raise
    except Exception:
        context = playwright.chromium.launch_persistent_context(
            SESSION_DIR,
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_https_errors=True,
        )

    diagnostic = PaymentMethodDiagnostic()
    page = context.new_page()
    context.on("response", diagnostic.capture)

    try:
        if restore_auth_state(context):
            print(f"Restored saved login state from: {AUTH_STATE_PATH}")
        print("Checking session...")
        ensure_authenticated(page, context)
        save_auth_state(context)

        if page.url == "about:blank":
            page.goto(P2P_REFERER)

        print("\nPAYMENT METHOD DIAGNOSTIC")
        print("This mode will not open or modify an Excel workbook.")
        print("In the opened browser, view the Overview section of a confirmed EFT invoice.")
        input("After Payment Method: EFT is visible, press Enter here...")
        page.wait_for_timeout(2500)

        print("\nNow view the Overview section of a confirmed cheque invoice.")
        input("After its Payment Method is visible, press Enter here...")
        page.wait_for_timeout(2500)
    finally:
        output_path = diagnostic.save()
        context.close()

    print(f"\nSanitized diagnostic report saved to:\n  {output_path}")
    print(f"Captured {len(diagnostic.records)} distinct JSON response shapes.")
    return output_path

# ---------------------------------------------------------------------------
# API calls
# ---------------------------------------------------------------------------

def lookup_ir_po(invoice, page, headers):
    """Returns (ir_id, po_id) as ints. Either can be None (contract = no PO)."""
    resp = page.request.post(
        IRGRID_URL,
        params={
            "NumRecords":      2,
            "search":          str(invoice),
            "enddate":         "",
            "invoicedatefrom": "",
            "invoicedateto":   "",
            "startdate":       "",
        },
        data=json.dumps([]),
        headers=headers,
    )
    records = resp.json().get("Records", [])
    if not records:
        return None, None

    invoice_key = normalize_invoice_key(invoice)
    matching_record = None
    for rec in records:
        for field in IRGRID_INVOICE_FIELDS:
            if normalize_invoice_key(rec.get(field)) == invoice_key:
                matching_record = rec
                break
        if matching_record:
            break

    if not matching_record:
        return None, None

    return (
        _first_int_field(matching_record, IRGRID_IR_FIELDS),
        _first_int_field(matching_record, IRGRID_PO_FIELDS),
    )


def search_irgrid_candidates(invoice, page, headers):
    """Return all exact-invoice candidates from a targeted IRGrid search."""
    response = page.request.post(
        IRGRID_URL,
        params={
            "NumRecords":      20,
            "search":          str(invoice),
            "enddate":         "",
            "invoicedatefrom": "",
            "invoicedateto":   "",
            "startdate":       "",
        },
        data=json.dumps([]),
        headers=headers,
    )
    invoice_key = normalize_invoice_key(invoice)
    matches = []
    for record in _unwrap_irgrid_records(response.json()):
        for field in IRGRID_INVOICE_FIELDS:
            if normalize_invoice_key(record.get(field)) == invoice_key:
                matches.append((record, field))
    return _deduplicate_invoice_matches(matches)


def normalize_invoice_key(value):
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.lower()


VENDOR_LEGAL_SUFFIXES = {
    "co", "company", "corp", "corporation", "inc", "incorporated",
    "limited", "ltd", "llc", "lp", "llp",
}


def normalize_vendor_name(value):
    """Normalize a contractor name without relying on exact punctuation."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.lower().replace("&", " and ")
    tokens = re.findall(r"[a-z0-9]+", text)
    tokens = [
        token for token in tokens
        if token not in VENDOR_LEGAL_SUFFIXES and token != "the"
    ]
    return " ".join(tokens)


def vendor_name_variants(value):
    """Return full/legal/operating-name variants for flexible comparison."""
    raw = str(value or "").strip()
    if not raw:
        return []
    segments = re.split(
        r"\b(?:operating\s+as|doing\s+business\s+as|trading\s+as)\b"
        r"|(?<![a-z])o\s*/\s*a(?![a-z])"
        r"|(?<![a-z])d\s*/\s*b\s*/\s*a(?![a-z])",
        raw,
        flags=re.IGNORECASE,
    )
    variants = []
    for candidate in [raw, *segments]:
        normalized = normalize_vendor_name(candidate)
        if normalized and normalized not in variants:
            variants.append(normalized)
    return variants


def vendor_name_score(left, right):
    """Return a conservative 0..1 similarity score for contractor names."""
    best = 0.0
    for left_name in vendor_name_variants(left):
        for right_name in vendor_name_variants(right):
            if left_name == right_name:
                return 1.0
            left_tokens = set(left_name.split())
            right_tokens = set(right_name.split())
            common = left_tokens & right_tokens
            sequence = SequenceMatcher(None, left_name, right_name).ratio()
            score = sequence
            if len(common) >= 2:
                containment = len(common) / min(len(left_tokens), len(right_tokens))
                jaccard = len(common) / len(left_tokens | right_tokens)
                score = max(score, 0.82 * containment + 0.18 * jaccard)
            elif len(common) == 1:
                # A single shared word such as "Plumbing" is not sufficient
                # evidence unless the normalized names are otherwise similar.
                score = min(score, 0.72)
            best = max(best, score)
    return best


def invoice_vendor_cache_key(invoice, contractor):
    return normalize_invoice_key(invoice), normalize_vendor_name(contractor)


def _first_int_field(record, fields):
    for field in fields:
        if field in record:
            value = to_int_id(record.get(field))
            if value is not None:
                return value
    return None


def _first_text_field(record, fields):
    for field in fields:
        if field in record and not is_empty(record.get(field)):
            return str(record.get(field)).strip()
    return ""


def _deduplicate_invoice_matches(matches):
    field_priority = {field: idx for idx, field in enumerate(IRGRID_INVOICE_FIELDS)}
    deduplicated = {}
    for record, matched_field in matches:
        if not isinstance(record, dict):
            continue
        ir_id = _first_int_field(record, IRGRID_IR_FIELDS)
        key = ("ir", ir_id) if ir_id is not None else ("object", id(record))
        existing = deduplicated.get(key)
        if existing is None or field_priority.get(
            matched_field, len(field_priority)
        ) < field_priority.get(existing[1], len(field_priority)):
            deduplicated[key] = (record, matched_field)
    return list(deduplicated.values())


def _money_value(value):
    if is_empty(value):
        return None
    text = str(value).strip().replace(",", "").replace("$", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return round(float(text), 2)
    except (TypeError, ValueError):
        return None


def _excel_date_key(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        pass
    for date_format in ("%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, date_format).date().isoformat()
        except ValueError:
            continue
    return None


def _yardi_date_key(value):
    parsed = _parse_date_value(value)
    return None if parsed == datetime.min else parsed.date().isoformat()


def _record_value(record, details, fields):
    value = _first_text_field(record, fields)
    if value:
        return value
    return _first_text_field(details or {}, fields)


def _secondary_candidate_evidence(record, details, row_context):
    """Return (matches, mismatches) for optional property/amount/date checks."""
    row_context = row_context or {}
    matches = 0
    mismatches = 0

    excel_property = normalize_vendor_name(row_context.get("property"))
    yardi_property = normalize_vendor_name(
        _record_value(record, details, IRGRID_PROPERTY_FIELDS)
    )
    if excel_property and yardi_property:
        if (excel_property == yardi_property
                or excel_property in yardi_property
                or yardi_property in excel_property):
            matches += 1
        else:
            mismatches += 1

    excel_amount = _money_value(row_context.get("amount"))
    yardi_amount = _money_value(_record_value(record, details, IRGRID_AMOUNT_FIELDS))
    if excel_amount is not None and yardi_amount is not None:
        if abs(excel_amount - yardi_amount) <= 0.01:
            matches += 1
        else:
            mismatches += 1

    excel_date = _excel_date_key(row_context.get("invoice_date"))
    yardi_date = _yardi_date_key(_record_value(record, details, IRGRID_DATE_FIELDS))
    if excel_date and yardi_date:
        if excel_date == yardi_date:
            matches += 1
        else:
            mismatches += 1

    return matches, mismatches


def select_irgrid_candidate(matches, contractor, detail_loader=None,
                            row_context=None):
    """Select one invoice record, or refuse an unsafe/ambiguous match."""
    candidates = _deduplicate_invoice_matches(matches)
    result = {
        "record": None,
        "matched_field": None,
        "resolution": "not-found",
        "candidate_count": len(candidates),
        "score": 0.0,
    }
    if not candidates:
        return result

    if len(candidates) == 1:
        record, matched_field = candidates[0]
        # Both lookup paths supply exact invoice matches. Contractor names are
        # needed only to distinguish multiple distinct invoice records.
        result.update({
            "record": record,
            "matched_field": matched_field,
            "resolution": "unique-invoice",
        })
        return result

    contractor_text = str(contractor or "").strip()
    if not contractor_text:
        result["resolution"] = "ambiguous"
        return result

    evaluated = []
    for record, matched_field in candidates:
        vendor = _first_text_field(record, IRGRID_VENDOR_FIELDS)
        score = vendor_name_score(contractor_text, vendor)
        details = {}
        source = "vendor-local"
        if not vendor and detail_loader:
            ir_id = _first_int_field(record, IRGRID_IR_FIELDS)
            details = detail_loader(ir_id) if ir_id else {}
            detail_vendor = _first_text_field(details or {}, IRGRID_VENDOR_FIELDS)
            detail_score = vendor_name_score(contractor_text, detail_vendor)
            if detail_score > score:
                score = detail_score
                source = "vendor-detail"
        secondary_matches, secondary_mismatches = _secondary_candidate_evidence(
            record, details, row_context
        )
        evaluated.append({
            "record": record,
            "matched_field": matched_field,
            "score": score,
            "source": source,
            "secondary_matches": secondary_matches,
            "secondary_mismatches": secondary_mismatches,
        })

    strong = sorted(
        (candidate for candidate in evaluated if candidate["score"] >= 0.78),
        key=lambda candidate: candidate["score"],
        reverse=True,
    )
    if not strong:
        result["resolution"] = "contractor-mismatch"
        return result
    if len(strong) == 1 or strong[0]["score"] - strong[1]["score"] >= 0.12:
        winner = strong[0]
    else:
        contenders = [
            candidate for candidate in strong
            if strong[0]["score"] - candidate["score"] < 0.12
        ]
        contenders.sort(
            key=lambda candidate: (
                candidate["secondary_matches"] - candidate["secondary_mismatches"],
                candidate["secondary_matches"],
            ),
            reverse=True,
        )
        best_secondary = (
            contenders[0]["secondary_matches"] - contenders[0]["secondary_mismatches"]
        )
        next_secondary = (
            contenders[1]["secondary_matches"] - contenders[1]["secondary_mismatches"]
            if len(contenders) > 1 else -999
        )
        if contenders[0]["secondary_matches"] < 1 or best_secondary <= next_secondary:
            result["resolution"] = "ambiguous"
            result["score"] = strong[0]["score"]
            return result
        winner = contenders[0]
        winner["source"] = "vendor-secondary"

    result.update({
        "record": winner["record"],
        "matched_field": winner["matched_field"],
        "resolution": winner["source"],
        "score": winner["score"],
    })
    return result


def _unwrap_irgrid_records(raw):
    if isinstance(raw, dict):
        records = raw.get("Records") or raw.get("records") or raw.get("Data") or raw.get("data")
        return records if isinstance(records, list) else []
    if isinstance(raw, list):
        return raw
    return []


def preload_irgrid(page, headers):
    """
    Pull the IRGrid once, matching the website's broad "Any" search.
    Returns a local invoice index so per-row IRGrid searches are avoided.
    """
    print("\nPreloading IRGrid records for local lookup...")
    resp = page.request.post(
        IRGRID_URL,
        params={
            "NumRecords":      BULK_IRGRID_RECORD_LIMIT,
            "search":          "",
            "enddate":         "",
            "invoicedatefrom": "",
            "invoicedateto":   "",
            "startdate":       "",
        },
        data=json.dumps([]),
        headers=headers,
    )
    records = _unwrap_irgrid_records(resp.json())
    index = {}

    for record in records:
        for field in IRGRID_INVOICE_FIELDS:
            if field not in record or is_empty(record.get(field)):
                continue
            key = normalize_invoice_key(record.get(field))
            if key:
                index.setdefault(key, []).append((record, field))

    print(f"  IRGrid records loaded : {len(records)}")
    print(f"  Local lookup keys     : {len(index)}")
    return index


def lookup_ir_po_from_bulk(invoice, irgrid_index, contractor="",
                           detail_loader=None, row_context=None):
    """Return IDs plus contractor-aware resolution details from bulk data."""
    matches = irgrid_index.get(normalize_invoice_key(invoice), [])
    selection = select_irgrid_candidate(
        matches,
        contractor,
        detail_loader=detail_loader,
        row_context=row_context,
    )
    record = selection.get("record")
    if record is None:
        return None, None, None, "", selection
    matched_field = selection["matched_field"]
    return (
        _first_int_field(record, IRGRID_IR_FIELDS),
        _first_int_field(record, IRGRID_PO_FIELDS),
        matched_field,
        _first_text_field(record, IRGRID_STATUS_FIELDS),
        selection,
    )


def _unwrap_grid_records(raw):
    if isinstance(raw, dict):
        records = raw.get("Records") or raw.get("records") or raw.get("Data") or raw.get("data")
        return records if isinstance(records, list) else []
    if isinstance(raw, list):
        return raw
    return []


def preload_pogrid(page, headers):
    """Pull the PO grid once for local PO-number and description lookup."""
    print("\nPreloading PO records for local lookup...")
    resp = page.request.post(
        POGRID_URL,
        params={
            "NumRecords":      BULK_PO_RECORD_LIMIT,
            "dateFrom":        "",
            "dateTo":          "",
            "showMarketplace": 0,
        },
        data=json.dumps([]),
        headers=headers,
    )
    records = _unwrap_grid_records(resp.json())
    index = {"by_po": {}, "records": records}

    for record in records:
        for field in POGRID_PO_FIELDS:
            if field not in record or is_empty(record.get(field)):
                continue
            key = normalize_invoice_key(record.get(field))
            if key:
                index["by_po"].setdefault(key, []).append((record, field))

    print(f"  PO records loaded     : {len(records)}")
    print(f"  Local PO lookup keys  : {len(index['by_po'])}")
    return index


def _description_contains_invoice(description, invoice):
    """Match an invoice as a complete description token, ignoring case."""
    invoice_text = str(invoice or "").strip()
    if invoice_text.endswith(".0"):
        invoice_text = invoice_text[:-2]
    if not invoice_text:
        return False
    pattern = rf"(?<![A-Za-z0-9]){re.escape(invoice_text)}(?![A-Za-z0-9])"
    return re.search(pattern, str(description or ""), flags=re.IGNORECASE) is not None


def _po_is_closed(record):
    """Recognize explicit header closure evidence; never infer it from notes/IDs."""
    status_fields = {"postatus", "status", "postatusname", "statusname",
                     "postatusdescription", "statusdescription", "postatustext", "statustext"}
    for field, value in record.items():
        key = re.sub(r"[^a-z0-9]", "", str(field).casefold())
        if key in {"isclosed", "poisclosed", "ispoclosed", "poclosed"}:
            if value is True or type(value) in (int, float) and value == 1:
                return True
            if isinstance(value, str) and value.strip().casefold() in {"true", "1"}:
                return True
        if key not in status_fields:
            continue
        values = [value]
        if isinstance(value, dict):
            values = [v for k, v in value.items()
                      if str(k).casefold() in {"name", "text", "label", "description"}]
        for label in values:
            if not isinstance(label, str):
                continue
            label = html.unescape(re.sub(r"<[^>]*>", " ", label))
            if " ".join(label.split()).casefold() == "closed":
                return True
    return False


def _lookup_po_from_description_records(invoice, records):
    """
    Return (po_id, matched_field, ambiguous_count) for an invoice appearing
    in a non-closed PO description. Return only an unambiguous match.
    """
    # One closed header excludes every grid row for that PO, including rows
    # whose projection omitted status. Never count a duplicate projection as active.
    closed_ids = {_first_int_field(r, POGRID_PO_FIELDS) for r in records if _po_is_closed(r)}
    matches = {}
    for record in records:
        if _po_is_closed(record) or _first_int_field(record, POGRID_PO_FIELDS) in closed_ids:
            continue
        for field in POGRID_DESCRIPTION_FIELDS:
            if not _description_contains_invoice(record.get(field), invoice):
                continue
            po_id = _first_int_field(record, POGRID_PO_FIELDS)
            if po_id is not None:
                matches.setdefault(po_id, field)

    if len(matches) != 1:
        return None, None, len(matches)
    po_id, matched_field = next(iter(matches.items()))
    return po_id, matched_field, 0


def _resolve_po_description_records(invoice, records, page, headers, closure_cache=None):
    result = _lookup_po_from_description_records(invoice, records)
    if result[2] < 2 or page is None:
        return result
    cache = closure_cache if closure_cache is not None else {}
    candidates = {_first_int_field(r, POGRID_PO_FIELDS) for r in records
                  if not _po_is_closed(r) and any(
                      _description_contains_invoice(r.get(f), invoice) for f in POGRID_DESCRIPTION_FIELDS)}
    for po_id in sorted(p for p in candidates if p is not None):
        if po_id in cache:
            continue
        try:
            response = page.request.get(f"{API_BASE}/po/{po_id}", headers=headers)
            if not response.ok:
                raise RuntimeError(f"HTTP {response.status}")
            detail = response.json()
            if not isinstance(detail, dict):
                raise RuntimeError("Unexpected PO header response")
            # Inspect the header only, never a closed line in DetailList.
            cache[po_id] = _po_is_closed(detail)
            print(f"  PO {po_id}: {'closed; excluded from invoice matching' if cache[po_id] else 'no closed header flag'}")
        except RequestStopped:
            raise
        except Exception as exc:
            # An unreadable candidate remains in the ambiguity set.
            print(f"  PO {po_id}: closure could not be confirmed ({type(exc).__name__}); candidate kept")
        pause_after_lookup(page)
    filtered = [r for r in records
                if not cache.get(_first_int_field(r, POGRID_PO_FIELDS), False)]
    return _lookup_po_from_description_records(invoice, filtered)


def lookup_po_from_description(invoice, pogrid_index, page=None, headers=None):
    """Resolve ambiguous grid matches against read-only PO headers once per run."""
    return _resolve_po_description_records(
        invoice, pogrid_index.get("records", []), page, headers or {},
        pogrid_index.setdefault("closure_cache", {}))


def search_po_by_description(invoice, page, headers):
    """Mirror the website PO search and verify the invoice in PODesc."""
    search_year = datetime.now().year
    resp = page.request.post(
        POGRID_URL,
        params={
            "dateFrom": f"01/01/{search_year}",
            "dateTo": f"12/31/{search_year}",
            "search": str(invoice),
        },
        data=json.dumps({}),
        headers=headers,
    )
    records = _unwrap_grid_records(resp.json())
    return _resolve_po_description_records(invoice, records, page, headers)


def _unwrap_milestones(raw):
    """Handle both bare list and wrapped {"data": [...]} response."""
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        return (
            raw.get("data") or raw.get("Records") or raw.get("Data")
            or next((v for v in raw.values() if isinstance(v, list)), [])
        )
    return []


def _parse_date_value(raw_date):
    """Parse a Yardi date while preserving its time for milestone ordering."""
    if not raw_date or str(raw_date) in ("None", ""):
        return datetime.min
    date_text = str(raw_date).strip()
    try:
        parsed = datetime.fromisoformat(date_text.replace("Z", "+00:00"))
    except ValueError:
        parsed = None
        # P2P displays slash-formatted dates as M/D/YYYY. Prefer that source
        # convention for ambiguous values such as 10/2/2026, then format the
        # result into the registry's D/M/YYYY convention.
        for date_format in ("%m/%d/%Y", "%d/%m/%Y"):
            try:
                parsed = datetime.strptime(date_text, date_format)
                break
            except ValueError:
                continue
        if parsed is None:
            return datetime.min

    if parsed.tzinfo is not None:
        # Preserve Yardi's displayed local date/time while making values
        # consistently comparable with timestamps that have no offset.
        parsed = parsed.replace(tzinfo=None)
    return parsed


def _format_milestone_date(date_str):
    """Convert a Yardi date/time to D/M/YYYY. Returns '' on failure."""
    dt = _parse_date_value(date_str)
    return "" if dt == datetime.min else f"{dt.day}/{dt.month}/{dt.year}"


def _parse_milestone_date(milestone, preferred_date_field):
    fields = (preferred_date_field,) + tuple(
        field for field in MILESTONE_DATE_FIELDS if field != preferred_date_field
    )
    for field in fields:
        parsed = _parse_date_value(milestone.get(field, ""))
        if parsed != datetime.min:
            return parsed
    return datetime.min


def _is_truthy_api_flag(value):
    return value is True or str(value).strip().lower() in ("1", "true", "yes")


def _has_parseable_date(milestone, field):
    return _parse_date_value(milestone.get(field, "")) != datetime.min


def _is_completed(milestone, date_field=None):
    if _is_truthy_api_flag(milestone.get("IsComplete")):
        return True
    if date_field and _has_parseable_date(milestone, date_field):
        return True
    return _parse_milestone_date(milestone, date_field or "Date") != datetime.min


def _looks_current(milestone):
    current_keys = (
        "IsInProgress", "InProgress", "IsCurrent", "Current", "IsCurrentStep",
        "CurrentStep", "IsActive", "Active", "IsPending", "Pending",
    )
    if any(_is_truthy_api_flag(milestone.get(key)) for key in current_keys):
        return True

    for key, value in milestone.items():
        key_text = str(key).lower()
        value_text = str(value).strip().lower()
        if ("current" in key_text or "progress" in key_text) and _is_truthy_api_flag(value):
            return True
        if key_text in ("status", "state", "class", "cssclass") and (
            "current" in value_text or "progress" in value_text or "pending" in value_text
        ):
            return True

    return False


def _seq_value(milestone):
    sequence_fields = (
        "Seq", "Sequence", "SequenceNo", "SeqNo", "StepSequence",
        "Order", "SortOrder", "StepOrder",
    )
    for field in sequence_fields:
        try:
            value = milestone.get(field)
            if value not in (None, ""):
                return float(value)
        except (ValueError, TypeError):
            continue
    return 0.0


def _completed_title_priority(milestone):
    """Provide a deterministic tie-breaker for common terminal workflow steps."""
    title = _milestone_title(milestone).strip().lower()
    if title == "paid" or "cheque" in title:
        return 500
    if "approved" in title or title == "approve":
        return 400
    if "completed" in title or "closed" in title:
        return 300
    if "invoice created" in title:
        return 200
    if "create po" in title:
        return 100
    if "po creation" in title:
        return 50
    return 0


def _milestone_sort_key(milestone, preferred_date_field, index=0):
    return (
        _parse_milestone_date(milestone, preferred_date_field),
        _seq_value(milestone),
        _completed_title_priority(milestone),
        -index,
    )


def _latest_milestone(milestones, preferred_date_field):
    return max(
        enumerate(milestones),
        key=lambda item: _milestone_sort_key(item[1], preferred_date_field, item[0]),
    )[1]


def _next_milestone(milestones, preferred_date_field):
    return min(
        enumerate(milestones),
        key=lambda item: (
            _parse_milestone_date(item[1], preferred_date_field),
            _seq_value(item[1]),
            item[0],
        ),
    )[1]


def _milestone_date(milestone, preferred_date_field):
    fields = (preferred_date_field,) + tuple(
        field for field in MILESTONE_DATE_FIELDS if field != preferred_date_field
    )
    for field in fields:
        formatted = _format_milestone_date(str(milestone.get(field, "")))
        if formatted:
            return formatted
    return ""


def _milestone_title(milestone):
    for field in ("Title", "Name", "StepName", "MilestoneName", "Description"):
        title = str(milestone.get(field, "")).strip()
        if title:
            return title
    return "Submitted"


def _current_milestone(milestones, date_field):
    """
    Return (title, formatted_date) for the current visible milestone.
    In-progress milestones are preferred because they reflect where the item
    currently sits in the workflow. If none are in progress, use the latest
    completed milestone.
    date_field is 'RevisedDate' for IR milestones, 'DtCompleted' for PO milestones.
    """
    current = [m for m in milestones if _looks_current(m)]
    if current:
        m = _latest_milestone(current, date_field)
        return _milestone_title(m), _milestone_date(m, date_field)

    completed = [m for m in milestones if _is_completed(m, date_field)]
    if completed:
        latest = _latest_milestone(completed, date_field)
        title  = _milestone_title(latest)
        date   = _milestone_date(latest, date_field)
        return title, date

    next_steps = [m for m in milestones if not _is_completed(m, date_field)]
    if not next_steps:
        return "Submitted", ""
    m = _next_milestone(next_steps, date_field)
    return _milestone_title(m), _milestone_date(m, date_field)


def _format_status(title, date, prefix=""):
    if title.strip().lower() == "paid" and date:
        status = f"Cheque has issued on {date}"
        return f"{prefix}{status}" if prefix else status
    status = f"{title}  {date}".strip() if date else title
    return f"{prefix}{status}" if prefix else status


def _normalize_payment_method_name(value):
    text = re.sub(r"[^a-z0-9]", "", str(value or "").lower())
    if text in ("eft", "electronicfundtransfer", "electronicfundstransfer"):
        return "eft"
    if text in ("check", "cheque"):
        return "check"
    return None


def is_na_payment_method(value):
    """Recognize an unknown-method marker; it is final only on paid rows."""
    return isinstance(value, str) and value.strip().casefold() == "n/a"


def _payment_method_values_equal(left, right):
    try:
        return float(left) == float(right)
    except (TypeError, ValueError):
        return str(left or "").strip().lower() == str(right or "").strip().lower()


def selected_payment_method_name(details):
    """Use the selected API value/option, never infer from payment references."""
    if not isinstance(details, dict):
        return None
    selected = details.get("PaymentMethod")
    if isinstance(selected, str) and selected.strip():
        try:
            float(selected)
        except ValueError:
            return selected.strip()
    options = details.get("PaymentMethodList")
    if not isinstance(options, list):
        return None
    names = ("name", "Name", "label", "Label", "text", "Text")
    values = ("id", "Id", "ID", "value", "Value", "code", "Code")
    flags = ("selected", "Selected", "isSelected", "IsSelected")
    def name(option):
        value = next((option[field] for field in names if field in option), None) if isinstance(option, dict) else option
        return value.strip() if isinstance(value, str) and value.strip() else None
    def is_selected(value):
        return value is True or value == 1 or (isinstance(value, str) and value.casefold() == "true")
    flagged = [name(option) for option in options if isinstance(option, dict) and any(is_selected(option.get(field)) for field in flags)]
    if flagged:
        return flagged[0] if len(flagged) == 1 else None
    if selected is None or isinstance(selected, bool) or (isinstance(selected, str) and not selected.strip()):
        return None
    explicit = any(isinstance(option, dict) and any(field in option for field in values) for option in options)
    if explicit:
        matched = [name(option) for option in options if isinstance(option, dict) and any(
            _payment_method_values_equal(selected, option[field]) for field in values if field in option)]
        return matched[0] if len(matched) == 1 else None
    try:
        number = float(selected)
        if not number.is_integer():
            return None
        index = int(number)
    except (ValueError, TypeError, OverflowError):
        return None
    return name(options[index]) if 0 <= index < len(options) else None


def resolve_payment_method(details):
    """Retain the EFT/check codes used by existing payment-status logic."""
    return _normalize_payment_method_name(selected_payment_method_name(details))


def payment_method_label(details):
    name = selected_payment_method_name(details)
    if not name:
        return None
    normalized = _normalize_payment_method_name(name)
    return {"eft": "EFT", "check": "Cheque"}.get(normalized, name)


def get_ir_details(ir_id, page, headers, detail_cache=None, stats=None):
    """Load one IR detail response, sharing it across matching and EFT logic."""
    if detail_cache is not None and ir_id in detail_cache:
        if stats is not None:
            stats["ir_detail_cache_hits"] += 1
        return detail_cache[ir_id]

    response = page.request.get(f"{IR_DETAIL_URL}/{ir_id}", headers=headers)
    if getattr(response, "ok", True) is False:
        raise RuntimeError("Invoice detail request failed")
    details = response.json()
    if not isinstance(details, dict):
        raise RuntimeError("Unexpected invoice detail response")
    if detail_cache is not None:
        detail_cache[ir_id] = details
    if stats is not None:
        stats["ir_detail_requests"] += 1
    return details


def lookup_ir_payment_method(ir_id, page, headers, detail_cache=None, stats=None):
    """Return 'eft', 'check', or None from the invoice Overview API."""
    details = get_ir_details(
        ir_id, page, headers, detail_cache=detail_cache, stats=stats
    )
    return resolve_payment_method(details)


def _format_ir_status(title, date, payment_method=None):
    if title.strip().lower() == "paid":
        if payment_method == "eft":
            return f"GL: EFT has initiated on {date}" if date else "GL: EFT has initiated"
        if payment_method != "check":
            method = f" ({payment_method})" if payment_method else ""
            return f"IR: Paid{method} on {date}" if date else f"IR: Paid{method}"
    return _format_status(title, date, prefix="IR: ")


def lookup_ir_milestone(ir_id, page, headers, detail_cache=None, stats=None, skip_payment_method=False):
    """
    Returns:
      status_str  : current visible P2P milestone status
      check_num   : str or None  (only when Paid milestone exists)
      check_date  : str D/M/YYYY or None
    """
    if detail_cache is None:
        detail_cache = {}
    resp = page.request.post(
        IR_MILESTONE_URL,
        params={"Id": ir_id},
        data=json.dumps([]),
        headers=headers,
    )
    milestones = _unwrap_milestones(resp.json())
    if not milestones:
        return "No milestones", None, None

    # Paid is terminal. Prefer its P2P values even if an earlier step is still
    # marked current by the API.
    paid_milestones = [
        m for m in milestones
        if _milestone_title(m).strip().lower() == "paid"
        and _is_completed(m, "RevisedDate")
    ]
    paid = (
        _latest_milestone(paid_milestones, "RevisedDate")
        if paid_milestones else None
    )

    payment_method = None
    if paid:
        title = "Paid"
        date = _milestone_date(paid, "RevisedDate")
        if not skip_payment_method:
            try:
                payment_method = lookup_ir_payment_method(
                    ir_id,
                    page,
                    headers,
                    detail_cache=detail_cache,
                    stats=stats,
                )
                if payment_method is None and ir_id in detail_cache:
                    payment_method = payment_method_label(detail_cache[ir_id])
            except RequestStopped:
                raise
            except Exception as exc:
                print(
                    f"payment method unavailable for IR {ir_id}; "
                    f"using generic paid status ({exc})"
                )
    else:
        title, date = _current_milestone(milestones, "RevisedDate")
    status_str = _format_ir_status(title, date, payment_method)

    # Also extract cheque data from the Paid milestone if present
    check_num  = None
    check_date = None
    if paid:
        check_num  = str(paid.get("CheckNum", "")).strip() or None
        check_date = _milestone_date(paid, "RevisedDate") or None

    return status_str, check_num, check_date


def lookup_po_milestone(po_id, page, headers):
    """Returns status_str for a PO-only row (no IR yet)."""
    resp = page.request.post(
        PO_MILESTONE_URL,
        params={"Id": po_id},
        data=json.dumps([]),
        headers=headers,
    )
    milestones = _unwrap_milestones(resp.json())
    if not milestones:
        return "PO: No milestones"
    title, date = _current_milestone(milestones, "DtCompleted")
    status_str  = _format_status(title, date, prefix="PO: ")
    return status_str

# ---------------------------------------------------------------------------
# Cell helpers
# ---------------------------------------------------------------------------

def is_empty(val):
    if val is None:
        return True
    return str(val).strip().lower() in ("", "nan", "none")


def to_int_id(val):
    if is_empty(val):
        return None
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


def is_payment_final_status(status):
    text = str(status or "").strip().lower()
    return ("cheque has issued" in text or "eft has initiated" in text
            or text.startswith("ir: paid"))


def is_cheque_issued(ch_num, ch_date, status):
    """Rows with issued cheque data are final and do not need P2P refresh."""
    if not is_empty(ch_num) and not is_empty(ch_date):
        return True
    status_text = str(status or "").strip().lower()
    return is_payment_final_status(status_text)


def has_user_entered_us_cheque_date(ch_date):
    """MM-DD-YYYY dates must be refreshed from P2P instead of trusted."""
    return bool(re.fullmatch(r"\d{1,2}-\d{1,2}-\d{4}", str(ch_date or "").strip()))


def get_col_map(ws):
    """Resolve current/legacy headers without moving or renaming any columns."""
    columns = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value is not None}
    def normalized(value):
        return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value))).casefold()
    aliases = {
        H_CH_DATE: (H_CH_DATE, "Ch. / EFT Date (DD/MM/YYYY)"),
        H_CH_NUM: (H_CH_NUM, "Ch.# / EFT No."),
        H_PAYMENT_METHOD: (H_PAYMENT_METHOD,),
    }
    for canonical in (H_STATUS, H_PO, H_IR, H_INVOICE, H_VENDOR, H_AMOUNT, H_INVOICE_DATE, H_PROPERTY, H_ACCOUNT):
        aliases[canonical] = (canonical,)
    for canonical, choices in aliases.items():
        matches = {cell.column for cell in ws[1] if cell.value is not None and normalized(cell.value) in {normalized(choice) for choice in choices}}
        if len(matches) > 1:
            raise ValueError(f"Sheet '{ws.title}' has duplicate headers for '{canonical}'. Keep only one matching column.")
        if matches:
            columns[canonical] = matches.pop()
    return columns


def payment_reference_value(value):
    """Keep leading zeros, long identifiers and alphanumeric EFT references."""
    text = str(value).strip()
    if re.fullmatch(r"[0-9]{1,15}", text) and (text == "0" or not text.startswith("0")):
        return int(text)
    return text



def registry_skip_reason(ws, row_idx, col_map=None):
    """Recognize explicit non-P2P markers; never infer from vendor or invoice text."""
    columns = get_col_map(ws) if col_map is None else col_map

    def text(header):
        column = columns.get(header)
        value = ws.cell(row=row_idx, column=column).value if column else None
        if not isinstance(value, str) or value.lstrip().startswith("="):
            return ""
        normalized = unicodedata.normalize("NFKC", value).casefold()
        normalized = re.sub(r"[\u2010-\u2015\u2212_-]+", " ", normalized)
        return re.sub(r"\s+", " ", normalized).strip()

    status, remarks = text(H_STATUS), text(H_REMARKS)
    cancelled = r"\bcancell?ed\b"
    # Remarks are job descriptions: require a cancellation annotation or an
    # explicit terminal marker instead of skipping an incidental mention.
    for candidate in (status,):
        if re.search(cancelled, candidate) and not re.search(r"\bnot\s+cancell?ed\b", candidate):
            return "Cancelled"
    if (re.search(r"\(\s*cancell?ed\s*\)", remarks)
            or re.fullmatch(r"(?:(?:invoice|payment|order|po)\s+)?cancell?ed[.!]?", remarks)):
        return "Cancelled"

    po_book = r"\bp\s*\.?\s*o\s*\.?\s+books?\b"
    preauthorized = r"\bpre\s*auth(?:ori[sz]|oi[sz])ed\s+(?:payments?|debits?)\b"
    for candidate in (status, remarks):
        if (re.search(po_book, candidate)
                and not re.search(r"\bnot\s+(?:(?:using|use|a)\s+)?" + po_book, candidate)):
            return "PO book"
        if (re.search(preauthorized, candidate)
                and not re.search(r"\b(?:not|never)\s+" + preauthorized, candidate)):
            return "Pre-authorized payment"
    return ""


SYNC_STAT_KEYS = (
    "filled_po", "filled_ir", "filled_cheque", "status_updated", "not_found",
    "excluded_rows",
    "bulk_hits", "fallback_lookups", "description_searches",
    "description_matches", "ambiguous_description_matches",
    "cheque_issued_skipped", "duplicate_invoice_groups",
    "vendor_matches_local", "vendor_matches_detail", "vendor_matches_secondary",
    "ambiguous_vendor_matches", "contractor_mismatches", "ir_detail_requests",
    "ir_detail_cache_hits", "ir_detail_failures",
    "payment_method_updated", "payment_method_unknown",
)


def new_sync_stats():
    return {key: 0 for key in SYNC_STAT_KEYS}


HISTORICAL_EFT_STAT_KEYS = (
    "completed_rows", "eft_updated", "eft_already_correct", "check_unchanged",
    "missing_ir", "unknown_payment_method", "eft_without_paid_milestone",
    "lookup_failures",
)


def new_historical_eft_stats():
    return {key: 0 for key in HISTORICAL_EFT_STAT_KEYS}


def is_eft_status(status):
    return str(status or "").strip().lower().startswith("gl: eft has initiated")


def refresh_historical_eft_sheet(
    ws, sheet_name, page, auth_headers, ir_detail_cache, ir_milestone_cache=None
):
    """Correct completed historical rows only when Yardi confirms EFT."""
    col_map = get_col_map(ws)
    col_status = col_map.get(H_STATUS)
    col_ch_date = col_map.get(H_CH_DATE)
    col_ch_num = col_map.get(H_CH_NUM)
    col_ir = col_map.get(H_IR)
    col_invoice = col_map.get(H_INVOICE)
    col_method = col_map.get(H_PAYMENT_METHOD)
    stats = new_historical_eft_stats()

    required = (col_status, col_ir, col_invoice)
    if not all(required):
        print(
            f"  [{sheet_name}] Required Status, IR, or Invoice column missing; "
            "sheet skipped."
        )
        return stats

    milestone_cache = ir_milestone_cache if ir_milestone_cache is not None else {}
    for row_idx in range(2, ws.max_row + 1):
        invoice_val = ws.cell(row=row_idx, column=col_invoice).value
        if is_empty(invoice_val):
            continue
        if col_method and is_na_payment_method(ws.cell(row=row_idx, column=col_method).value):
            continue

        status_val = ws.cell(row=row_idx, column=col_status).value
        ch_val = ws.cell(row=row_idx, column=col_ch_num).value if col_ch_num else None
        chd_val = (
            ws.cell(row=row_idx, column=col_ch_date).value if col_ch_date else None
        )
        if not is_cheque_issued(ch_val, chd_val, status_val):
            continue

        stats["completed_rows"] += 1
        if is_eft_status(status_val):
            stats["eft_already_correct"] += 1
            continue

        ir_id = to_int_id(ws.cell(row=row_idx, column=col_ir).value)
        if not ir_id:
            stats["missing_ir"] += 1
            continue

        label = f"  [{sheet_name}] Row {row_idx:>4} | {str(invoice_val).strip():<18}"
        try:
            was_cached = ir_id in ir_detail_cache
            payment_method = lookup_ir_payment_method(
                ir_id,
                page,
                auth_headers,
                detail_cache=ir_detail_cache,
            )
            if not was_cached:
                time.sleep(0.1)
        except RequestStopped:
            raise
        except Exception as exc:
            stats["lookup_failures"] += 1
            print(f"{label}  payment method lookup failed: {exc}")
            continue

        if payment_method == "check":
            stats["check_unchanged"] += 1
            continue
        if payment_method != "eft":
            stats["unknown_payment_method"] += 1
            print(f"{label}  payment method unknown; unchanged")
            continue

        if ir_id not in milestone_cache:
            try:
                milestone_cache[ir_id] = lookup_ir_milestone(
                    ir_id,
                    page,
                    auth_headers,
                    detail_cache=ir_detail_cache,
                )
                pause_after_lookup(page)
            except RequestStopped:
                raise
            except Exception as exc:
                stats["lookup_failures"] += 1
                print(f"{label}  milestone lookup failed: {exc}")
                continue

        status_str, _check_num, _check_date = milestone_cache[ir_id]
        if not is_eft_status(status_str):
            stats["eft_without_paid_milestone"] += 1
            print(f"{label}  EFT confirmed but no completed Paid milestone; unchanged")
            continue

        if status_val != status_str:
            ws.cell(row=row_idx, column=col_status).value = status_str
            stats["eft_updated"] += 1
            print(f"{label}  {status_str}")
        else:
            stats["eft_already_correct"] += 1

    return stats


def pause_after_lookup(page):
    """Desktop requests enforce the pause globally, including detail reads."""
    if getattr(getattr(page, "request", None), "uses_global_pacing", False) is not True:
        time.sleep(0.3)


def sync_sheet(ws, sheet_name, page, auth_headers, irgrid_index, pogrid_index,
               ir_cache, cheque_cache, po_status_cache, ir_detail_cache, progress_callback=None, attention_callback=None):
    """
    Process one worksheet. Fills P2P Status only when the column exists,
    and fills empty PO, IR, Ch.#, Ch. Date cells.
    Returns stats dict.
    """
    col_map = get_col_map(ws)

    col_status  = col_map.get(H_STATUS)
    col_ch_date = col_map.get(H_CH_DATE)
    col_ch_num  = col_map.get(H_CH_NUM)
    col_po      = col_map.get(H_PO)
    col_ir      = col_map.get(H_IR)
    col_invoice = col_map.get(H_INVOICE)
    col_vendor  = col_map.get(H_VENDOR)
    col_amount  = col_map.get(H_AMOUNT)
    col_invoice_date = col_map.get(H_INVOICE_DATE)
    col_property = col_map.get(H_PROPERTY)
    col_method = col_map.get(H_PAYMENT_METHOD)

    if not col_status:
        print(f"  '{H_STATUS}' column not found - status updates skipped.")

    if not col_invoice:
        print(f"  '{H_INVOICE}' column not found — skipping sheet.")
        return new_sync_stats()

    stats = new_sync_stats()
    duplicate_invoice_keys = set()

    for row_idx in range(2, ws.max_row + 1):
        invoice_val = ws.cell(row=row_idx, column=col_invoice).value
        if is_empty(invoice_val):
            continue   # blank / spacing row — skip silently

        if progress_callback is not None:
            progress_callback(row_idx)

        if registry_skip_reason(ws, row_idx, col_map):
            stats["excluded_rows"] += 1
            continue

        invoice = str(invoice_val).strip()
        label   = f"  [{sheet_name}] Row {row_idx:>4} | {invoice:<18}"
        contractor = (
            ws.cell(row=row_idx, column=col_vendor).value if col_vendor else ""
        )
        row_context = {
            "property": ws.cell(row=row_idx, column=col_property).value
            if col_property else "",
            "amount": ws.cell(row=row_idx, column=col_amount).value
            if col_amount else None,
            "invoice_date": ws.cell(row=row_idx, column=col_invoice_date).value
            if col_invoice_date else None,
        }

        def report_attention(reason):
            if attention_callback is not None:
                attention_callback(row_idx, reason)

        def load_candidate_details(candidate_ir_id):
            try:
                return get_ir_details(
                    candidate_ir_id,
                    page,
                    auth_headers,
                    detail_cache=ir_detail_cache,
                    stats=stats,
                )
            except RequestStopped:
                raise
            except Exception as exc:
                stats["ir_detail_failures"] += 1
                report_attention("Invoice details could not be retrieved from Yardi. Try syncing again.")
                print(f"{label}  IR detail lookup failed: {exc}")
                return {}

        def fill_payment_method(ir_id, final=False):
            if not col_method:
                return
            cell = ws.cell(row=row_idx, column=col_method)
            if cell.data_type == "f" or (final and is_na_payment_method(cell.value)
                    and is_payment_final_status(status_val)):
                return  # Preserve formulas and final unknown-payment decisions.
            if not ir_id and not final:
                return  # PO-stage rows never request or acquire a payment method.
            try:
                method = payment_method_label(get_ir_details(ir_id, page, auth_headers,
                    detail_cache=ir_detail_cache, stats=stats)) if ir_id else None
            except RequestStopped:
                raise
            except Exception:
                stats["ir_detail_failures"] += 1
                report_attention("The payment method could not be retrieved from Yardi. The existing value was kept; try syncing again.")
                return
            if method is None:
                stats["payment_method_unknown"] += 1
                if not final:
                    # An IR's selected method may change before payment is issued.
                    if is_na_payment_method(cell.value):
                        cell.value = None
                        stats["payment_method_updated"] += 1
                    return
                method = "N/A"
            if cell.value != method:
                cell.value = method
                cell.data_type = "s"
                stats["payment_method_updated"] += 1

        # Account Code is intentionally manual. Never collapse split invoice rows
        # or add detail requests solely to check coding (rollback 2026-11-08).
        # Read current cell values
        po_val  = ws.cell(row=row_idx, column=col_po).value  if col_po  else None
        ir_val  = ws.cell(row=row_idx, column=col_ir).value  if col_ir  else None
        ch_val  = ws.cell(row=row_idx, column=col_ch_num).value if col_ch_num else None
        chd_val = ws.cell(row=row_idx, column=col_ch_date).value if col_ch_date else None
        status_val = ws.cell(row=row_idx, column=col_status).value if col_status else None

        force_cheque_refresh = has_user_entered_us_cheque_date(chd_val)

        final_payment = is_payment_final_status(status_val)
        # Old versions wrote N/A even before an IR existed. Undo that premature
        # marker without touching formulas or intentional non-P2P skipped rows.
        if col_method and not final_payment and not to_int_id(ir_val):
            method_cell = ws.cell(row=row_idx, column=col_method)
            if method_cell.data_type != "f" and is_na_payment_method(method_cell.value):
                method_cell.value = None
                stats["payment_method_updated"] += 1

        if (not force_cheque_refresh
                and is_cheque_issued(ch_val, chd_val, status_val)
                and (final_payment or not to_int_id(ir_val))):
            # Backfill the new column even when paid rows otherwise stay final.
            if col_method and is_empty(ws.cell(row=row_idx, column=col_method).value):
                fill_payment_method(to_int_id(ir_val), final=final_payment)
            stats["cheque_issued_skipped"] += 1
            continue

        need_po  = is_empty(po_val)
        need_ir  = is_empty(ir_val)
        need_ch  = force_cheque_refresh or is_empty(ch_val)
        need_chd = force_cheque_refresh or is_empty(chd_val)

        current_ir_id = to_int_id(ir_val)
        current_po_id = to_int_id(po_val)

        # ── Step 1: IRGrid lookup if PO or IR is missing ────────────
        if need_po or need_ir:
            cache_key = invoice_vendor_cache_key(invoice, contractor)
            if cache_key not in ir_cache:
                ir_id, po_id, matched_field, bulk_status, selection = (
                    lookup_ir_po_from_bulk(
                        invoice,
                        irgrid_index,
                        contractor=contractor,
                        detail_loader=load_candidate_details,
                        row_context=row_context,
                    )
                )
                if selection["candidate_count"] > 1:
                    duplicate_invoice_keys.add(normalize_invoice_key(invoice))
                if matched_field:
                    resolution = selection["resolution"]
                    if resolution == "vendor-local":
                        stats["vendor_matches_local"] += 1
                    elif resolution == "vendor-detail":
                        stats["vendor_matches_detail"] += 1
                    elif resolution == "vendor-secondary":
                        stats["vendor_matches_secondary"] += 1
                    ir_cache[cache_key] = (
                        ir_id,
                        po_id,
                        f"bulk:{matched_field}:{resolution}",
                        bulk_status,
                    )
                elif selection["resolution"] in ("ambiguous", "contractor-mismatch"):
                    resolution = selection["resolution"]
                    if resolution == "ambiguous":
                        stats["ambiguous_vendor_matches"] += 1
                    else:
                        stats["contractor_mismatches"] += 1
                    ir_cache[cache_key] = (
                        None, None, f"vendor:{resolution}", ""
                    )
                else:
                    print(f"{label}  IRGrid fallback ...", end="  ", flush=True)
                    stats["fallback_lookups"] += 1
                    try:
                        fallback_matches = search_irgrid_candidates(
                            invoice, page, auth_headers
                        )
                        fallback_selection = select_irgrid_candidate(
                            fallback_matches,
                            contractor,
                            detail_loader=load_candidate_details,
                            row_context=row_context,
                        )
                        if fallback_selection["candidate_count"] > 1:
                            duplicate_invoice_keys.add(normalize_invoice_key(invoice))
                        fallback_record = fallback_selection["record"]
                        if fallback_record is not None:
                            ir_id = _first_int_field(
                                fallback_record, IRGRID_IR_FIELDS
                            )
                            po_id = _first_int_field(
                                fallback_record, IRGRID_PO_FIELDS
                            )
                            resolution = fallback_selection["resolution"]
                            if resolution == "vendor-local":
                                stats["vendor_matches_local"] += 1
                            elif resolution == "vendor-detail":
                                stats["vendor_matches_detail"] += 1
                            elif resolution == "vendor-secondary":
                                stats["vendor_matches_secondary"] += 1
                            ir_cache[cache_key] = (
                                ir_id, po_id, f"fallback:{resolution}", ""
                            )
                        elif fallback_selection["resolution"] in (
                            "ambiguous", "contractor-mismatch"
                        ):
                            resolution = fallback_selection["resolution"]
                            if resolution == "ambiguous":
                                stats["ambiguous_vendor_matches"] += 1
                            else:
                                stats["contractor_mismatches"] += 1
                            ir_cache[cache_key] = (
                                None, None, f"vendor:{resolution}", ""
                            )
                        else:
                            ir_cache[cache_key] = (None, None, "fallback", "")
                        pause_after_lookup(page)
                    except RequestStopped:
                        raise
                    except Exception as e:
                        print(f"error: {e}")
                        report_attention("The invoice search failed. Try syncing again before checking the invoice manually.")
                        ir_cache[cache_key] = (None, None, "error", "")

            ir_id, po_id, lookup_source, bulk_status = ir_cache[cache_key]
            if lookup_source.startswith("bulk:"):
                stats["bulk_hits"] += 1

            if (ir_id is None and po_id is None
                    and not lookup_source.startswith(("description:", "vendor:"))):
                po_id, description_field, ambiguous_count = lookup_po_from_description(
                    invoice, pogrid_index, page, auth_headers
                )
                if po_id is None and not ambiguous_count:
                    stats["description_searches"] += 1
                    try:
                        po_id, description_field, ambiguous_count = search_po_by_description(
                            invoice, page, auth_headers
                        )
                        pause_after_lookup(page)
                    except RequestStopped:
                        raise
                    except Exception as e:
                        report_attention("The PO description search failed. Try syncing again.")
                        print(f"{label}  PO description search error: {e}")

                if po_id is not None:
                    lookup_source = f"PO description:{description_field}"
                    stats["description_matches"] += 1
                elif ambiguous_count:
                    stats["ambiguous_description_matches"] += 1
                    print(
                        f"{label}  ambiguous PO description match "
                        f"({ambiguous_count} POs)"
                    )
                    lookup_source = "description:ambiguous"
                else:
                    lookup_source = "description:not-found"

                ir_cache[cache_key] = (ir_id, po_id, lookup_source, bulk_status)

            if ir_id is None and po_id is None:
                if lookup_source == "description:ambiguous":
                    report_attention("The invoice number matches multiple PO descriptions. Check which PO belongs to this invoice.")
                if lookup_source == "vendor:ambiguous":
                    print(f"{label}  ambiguous P2P match - contractor check required")
                    status_message = "Ambiguous P2P match - check contractor"
                    report_attention("Multiple Yardi invoices match this invoice number and contractor. Check the correct record.")
                elif lookup_source == "vendor:contractor-mismatch":
                    print(f"{label}  P2P contractor does not match registry")
                    status_message = "P2P contractor mismatch - check record"
                    report_attention("The Yardi contractor does not match the contractor in this registry row.")
                elif lookup_source == "fallback":
                    print("not found in P2P")
                    status_message = "Not in P2P"
                else:
                    print(f"{label}  not found in P2P")
                    status_message = "Not in P2P"
                if not lookup_source.startswith("vendor:"):
                    stats["not_found"] += 1
                    if lookup_source != "description:ambiguous":
                        report_attention("No matching invoice register or PO was found by invoice number. Check whether it has been entered in Yardi.")
                # Still try PO milestone if PO is manually entered
                if current_po_id:
                    pass   # handled in status step below
                else:
                    if col_status:
                        ws.cell(row=row_idx, column=col_status).value = status_message
                        stats["status_updated"] += 1
                    continue
            else:
                parts = []
                if need_po:
                    if po_id is not None:
                        ws.cell(row=row_idx, column=col_po).value = po_id
                        stats["filled_po"] += 1
                        current_po_id = po_id
                        parts.append(f"PO={po_id}")
                    else:
                        parts.append("PO=N/A(contract)")
                if need_ir and ir_id is not None:
                    ws.cell(row=row_idx, column=col_ir).value = ir_id
                    stats["filled_ir"] += 1
                    current_ir_id = ir_id
                    parts.append(f"IR={ir_id}")
                if parts:
                    print(f"{label}  {', '.join(parts)}  ({lookup_source})")


        # ── Step 2: Milestone lookup — IR path (preferred) ──────────

        need_ir_milestone = True

        if current_ir_id and need_ir_milestone:
            skip_method = bool(final_payment and col_method and is_na_payment_method(ws.cell(row=row_idx, column=col_method).value))
            milestone_key = (current_ir_id, "N/A") if skip_method else current_ir_id
            if milestone_key not in cheque_cache:
                print(f"{label}  IR milestones ...", end="  ", flush=True)
                try:
                    status_str, chk_num, chk_date = lookup_ir_milestone(
                        current_ir_id,
                        page,
                        auth_headers,
                        detail_cache=ir_detail_cache,
                        stats=stats,
                        skip_payment_method=skip_method,
                    )
                    cheque_cache[milestone_key] = (status_str, chk_num, chk_date)
                    pause_after_lookup(page)
                except RequestStopped:
                    raise
                except Exception as e:
                    print(f"error: {e}")
                    cheque_cache[milestone_key] = ("Error", None, None)

            status_str, chk_num, chk_date = cheque_cache[milestone_key]
            fill_payment_method(current_ir_id, final=is_payment_final_status(status_str))
            if status_str == "Error":
                report_attention("Yardi invoice workflow or payment details could not be retrieved. Try syncing again.")
            elif status_str == "No milestones":
                report_attention("Yardi returned no workflow history for this invoice register. Check its workflow in Yardi.")

            if force_cheque_refresh and (not chk_num or not chk_date):
                report_attention("Yardi has incomplete paid-cheque details. Existing cheque values were kept; check the cheque number and date.")
                print(
                    f"{label}  P2P Paid cheque data incomplete; "
                    "existing cheque values kept"
                )
                continue

            if col_status:
                ws.cell(row=row_idx, column=col_status).value = status_str
                stats["status_updated"] += 1

            # Fill cheque fields only if empty
            if need_ch and chk_num and col_ch_num:
                reference = payment_reference_value(chk_num)
                payment_cell = ws.cell(row=row_idx, column=col_ch_num)
                payment_cell.value = reference
                if isinstance(reference, str):
                    payment_cell.data_type = "s"
                stats["filled_cheque"] += 1

            if need_chd and chk_date and col_ch_date:
                ws.cell(row=row_idx, column=col_ch_date).value = chk_date

            print(f"{label}  {status_str}")

        # ── Step 3: PO milestone path (no IR yet) ───────────────────
        elif not current_ir_id and current_po_id:
            if current_po_id not in po_status_cache:
                print(f"{label}  PO milestones ...", end="  ", flush=True)
                try:
                    po_status = lookup_po_milestone(current_po_id, page, auth_headers)
                    po_status_cache[current_po_id] = po_status
                    pause_after_lookup(page)
                except RequestStopped:
                    raise
                except Exception as e:
                    print(f"error: {e}")
                    po_status_cache[current_po_id] = "Error"

            po_status = po_status_cache[current_po_id]
            if po_status == "Error":
                report_attention("Yardi PO workflow details could not be retrieved. Try syncing again.")
            elif "No milestones" in po_status:
                report_attention("Yardi returned no workflow history for this PO. Check its workflow in Yardi.")
            if col_status:
                ws.cell(row=row_idx, column=col_status).value = po_status
                stats["status_updated"] += 1
            print(f"{label}  {po_status}")

    stats["duplicate_invoice_groups"] = len(duplicate_invoice_keys)
    return stats

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run_historical_eft_refresh(wb, page, auth_headers):
    """Run the explicit one-time EFT correction across the workbook."""
    total = new_historical_eft_stats()
    ir_detail_cache = {}
    ir_milestone_cache = {}

    print("\nOne-time historical EFT status refresh")
    print("Only Yardi-confirmed EFT rows will be changed.")
    for sheet_name in wb.sheetnames:
        print(f"\n{'='*60}")
        print(f"  Sheet: {sheet_name}")
        print(f"{'='*60}")
        stats = refresh_historical_eft_sheet(
            wb[sheet_name],
            sheet_name,
            page,
            auth_headers,
            ir_detail_cache,
            ir_milestone_cache,
        )
        for key in total:
            total[key] += stats.get(key, 0)
        print(
            "  Sheet summary - "
            f"Completed reviewed: {stats['completed_rows']}  "
            f"EFT updated: {stats['eft_updated']}  "
            f"Already EFT: {stats['eft_already_correct']}  "
            f"Checks unchanged: {stats['check_unchanged']}  "
            f"Unknown: {stats['unknown_payment_method']}  "
            f"Failures: {stats['lookup_failures']}"
        )
    return total


def print_historical_eft_summary(total, registry_path, backup_path=None):
    print(f"\n{'='*60}")
    print("  HISTORICAL EFT REFRESH SUMMARY")
    print(f"{'='*60}")
    print(f"  Completed rows reviewed : {total['completed_rows']}")
    print(f"  EFT statuses updated    : {total['eft_updated']}")
    print(f"  EFT already correct     : {total['eft_already_correct']}")
    print(f"  Cheque rows unchanged   : {total['check_unchanged']}")
    print(f"  Missing IR number       : {total['missing_ir']}")
    print(f"  Unknown payment method  : {total['unknown_payment_method']}")
    print(f"  EFT without Paid step   : {total['eft_without_paid_milestone']}")
    print(f"  Lookup failures         : {total['lookup_failures']}")
    if backup_path:
        print(f"\n  Backup saved            : {backup_path}")
        print(f"  Workbook updated        : {registry_path}")
    else:
        print("\n  No confirmed status changes; workbook was not saved.")
    print(f"{'='*60}")


def main():
    if is_payment_method_diagnostic():
        with sync_playwright() as playwright:
            run_payment_method_diagnostic(playwright)
        return

    registry_path = get_registry_path()

    if not registry_path.lower().endswith((".xlsx", ".xlsm")):
        print(f"Error: selected file is not an Excel workbook:\n  {registry_path}")
        return

    if not os.path.isfile(registry_path):
        print(f"Error: registry file not found:\n  {registry_path}")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"Loading : {os.path.basename(registry_path)}")
    wb = load_workbook(registry_path)

    os.makedirs(SESSION_DIR, exist_ok=True)

    with sync_playwright() as p:
        try:
            context = p.chromium.launch_persistent_context(
                SESSION_DIR,
                channel="chrome",
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
                ignore_https_errors=True,
            )
        except RequestStopped:
            raise
        except Exception:
            context = p.chromium.launch_persistent_context(
                SESSION_DIR,
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
                ignore_https_errors=True,
            )

        if restore_auth_state(context):
            print(f"Restored saved login state from: {AUTH_STATE_PATH}")

        page = context.new_page()
        print("Checking session...")
        auth_headers = ensure_authenticated(page, context)
        if save_auth_state(context):
            print(f"Saved login state to: {AUTH_STATE_PATH}")

        try:
            page.evaluate("() => { window.moveTo(-10000, 0); }")
        except RequestStopped:
            raise
        except Exception:
            pass

        if is_historical_eft_refresh():
            total = run_historical_eft_refresh(wb, page, auth_headers)
            context.close()
            backup_path = None
            if total["eft_updated"]:
                backup_path = backup_existing_workbook(registry_path)
                wb.save(registry_path)
            print_historical_eft_summary(total, registry_path, backup_path)
            return

        # Shared caches across all sheets — avoids duplicate API calls
        ir_cache        = {}   # invoice_str  → (ir_id, po_id)
        cheque_cache    = {}   # ir_id        → (status_str, chk_num, chk_date)
        po_status_cache = {}   # po_id        → status_str
        ir_detail_cache = {}   # ir_id        → full /ir/{id} response

        print("\nP2P status: detailed milestone lookup")
        irgrid_index = preload_irgrid(page, auth_headers)
        pogrid_index = preload_pogrid(page, auth_headers)

        total = new_sync_stats()

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            print(f"\n{'='*60}")
            print(f"  Sheet: {sheet_name}")
            print(f"{'='*60}")

            stats = sync_sheet(ws, sheet_name, page, auth_headers,
                               irgrid_index, pogrid_index,
                               ir_cache, cheque_cache, po_status_cache,
                               ir_detail_cache)

            for k in total:
                total[k] += stats.get(k, 0)

            print(f"\n  Sheet summary — "
                  f"Status updated: {stats['status_updated']}  "
                  f"PO filled: {stats['filled_po']}  "
                  f"IR filled: {stats['filled_ir']}  "
                  f"Cheque filled: {stats['filled_cheque']}  "
                  f"Cheque issued skipped: {stats['cheque_issued_skipped']}  "
                  f"Not in P2P: {stats['not_found']}  "
                  f"PO description searches: {stats['description_searches']}  "
                  f"PO description matches: {stats['description_matches']}  "
                  f"Ambiguous descriptions: {stats['ambiguous_description_matches']}  "
                  f"Duplicate invoices: {stats['duplicate_invoice_groups']}  "
                  f"Vendor local/detail/secondary: "
                  f"{stats['vendor_matches_local']}/"
                  f"{stats['vendor_matches_detail']}/"
                  f"{stats['vendor_matches_secondary']}  "
                  f"Vendor ambiguous/mismatch: "
                  f"{stats['ambiguous_vendor_matches']}/"
                  f"{stats['contractor_mismatches']}  "
                  f"IR detail requests/cache hits: "
                  f"{stats['ir_detail_requests']}/"
                  f"{stats['ir_detail_cache_hits']}  "
                  f"Bulk hits: {stats['bulk_hits']}  "
                  f"Fallback searches: {stats['fallback_lookups']}")

        context.close()

    backup_path = backup_existing_workbook(registry_path)
    wb.save(registry_path)

    print(f"\n{'='*60}")
    print("  TOTAL SUMMARY")
    print(f"{'='*60}")
    print(f"  Status updated  : {total['status_updated']}")
    print(f"  PO filled       : {total['filled_po']}")
    print(f"  IR filled       : {total['filled_ir']}")
    print(f"  Cheque filled   : {total['filled_cheque']}")
    print(f"  Cheque issued skipped: {total['cheque_issued_skipped']}")
    print(f"  Not in P2P yet  : {total['not_found']}")
    print(f"  PO description searches: {total['description_searches']}")
    print(f"  PO description matches: {total['description_matches']}")
    print(f"  Ambiguous descriptions: {total['ambiguous_description_matches']}")
    print(f"  Duplicate invoice groups: {total['duplicate_invoice_groups']}")
    print(f"  Vendor matches from bulk: {total['vendor_matches_local']}")
    print(f"  Vendor matches from detail: {total['vendor_matches_detail']}")
    print(f"  Vendor matches by secondary fields: {total['vendor_matches_secondary']}")
    print(f"  Ambiguous vendor matches: {total['ambiguous_vendor_matches']}")
    print(f"  Contractor mismatches: {total['contractor_mismatches']}")
    print(f"  IR detail requests: {total['ir_detail_requests']}")
    print(f"  IR detail cache hits: {total['ir_detail_cache_hits']}")
    print(f"  IR detail failures: {total['ir_detail_failures']}")
    print(f"  Bulk IRGrid hits: {total['bulk_hits']}")
    print(f"  Fallback searches: {total['fallback_lookups']}")
    print(f"\n  Backup saved    : {backup_path}")
    print(f"  Workbook updated: {registry_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    raise SystemExit("Public showcase: live worker execution is disabled.")
    try:
        main()
    except RequestStopped:
        raise
    except Exception as e:
        import traceback
        print(f"\nUnexpected error: {e}")
        traceback.print_exc()
    input("\nPress Enter to close...")

