"""Pacing and blocked-run tests use fake time and mocked/local endpoints only."""
from pathlib import Path
import contextlib
import json
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "PropertyManager.App" / "Backend"))
import request_policy as policy
import P2P_Registry_Sync as core
import registry_worker as worker
import session_transport
from openpyxl import Workbook


class Clock:
    def __init__(self):
        self.now = 1000.0

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


def response(status=200, body=None, headers=None, text=""):
    result = MagicMock()
    result.status = status
    result.ok = 200 <= status < 300
    result.headers = headers or {"content-type": "application/json"}
    result.json.return_value = [] if body is None else body
    result.text.return_value = text
    return result


class RequestPolicyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.clock = Clock()
        self.cooldown = policy.Cooldown(Path(self.temp.name) / "cooldown.json", self.clock)
        self.client = MagicMock()
        self.starts = []
        self.reply = response()

        def send(*args, **kwargs):
            self.starts.append(self.clock())
            self.clock.sleep(0.2)
            return self.reply
        self.client.get.side_effect = self.client.post.side_effect = send
        self.sync = SimpleNamespace(IRGRID_URL="https://fixture.test/api/IRGrid",
            POGRID_URL="https://fixture.test/api/POGrid", IR_DETAIL_URL="https://fixture.test/api/ir",
            IR_MILESTONE_URL="https://fixture.test/api/ir/milestones", PO_MILESTONE_URL="https://fixture.test/api/po/milestones")
        self.cancel = MagicMock()
        self.requests = policy.MeasuredRequests(self.client, self.sync, self.cooldown, self.cancel,
            clock=self.clock, sleep=self.clock.sleep)

    def test_all_endpoint_types_share_quiet_interval_and_metrics(self):
        self.requests.post(self.sync.IRGRID_URL, params={"search": ""}, data="[]")
        self.requests.get(self.sync.IR_DETAIL_URL + "/123")
        self.requests.post(self.sync.IR_MILESTONE_URL, params={"Id": 123})
        self.assertAlmostEqual(self.starts[1] - self.starts[0], 0.5)
        self.assertAlmostEqual(self.starts[2] - self.starts[1], 0.5)
        summary = self.requests.summary()
        self.assertEqual(summary["requests"], 3)
        self.assertAlmostEqual(summary["network_seconds"], 0.6)
        self.assertAlmostEqual(summary["pacing_seconds"], 0.6)
        self.assertEqual(summary["endpoints"]["IR details"]["p95_ms"], 200)
        self.assertNotIn("123", json.dumps(summary))
        self.assertNotIn("https", json.dumps(summary))

    def test_only_identical_targeted_searches_are_reused(self):
        self.requests.post(self.sync.IRGRID_URL, params={"search": "invoice-private", "NumRecords": 20}, data="[]")
        self.requests.post(self.sync.IRGRID_URL, params={"NumRecords": 20, "search": "invoice-private"}, data="[]")
        self.assertEqual(len(self.starts), 1)
        self.assertEqual(self.requests.summary()["reused_searches"], 1)
        self.assertNotIn("invoice-private", json.dumps(self.requests.summary()))
        # Bulk requests and changed filters are deliberately not reused.
        self.requests.post(self.sync.IRGRID_URL, params={"search": "invoice-private", "NumRecords": 21}, data="[]")
        self.requests.post(self.sync.IRGRID_URL, params={"search": ""}, data="[]")
        self.requests.post(self.sync.IRGRID_URL, params={"search": ""}, data="[]")
        self.assertEqual(len(self.starts), 4)

    def test_throttle_stops_without_retry_and_persists_server_wait(self):
        self.reply = response(429, headers={"Retry-After": "120"})
        with self.assertRaisesRegex(policy.RequestStopped, "120 seconds"):
            self.requests.get(self.sync.IR_DETAIL_URL + "/1")
        with self.assertRaises(policy.RequestStopped):
            self.requests.get(self.sync.IR_DETAIL_URL + "/2")
        self.assertEqual(len(self.starts), 1)
        another = policy.Cooldown(self.cooldown.path, self.clock)
        with self.assertRaisesRegex(policy.RequestStopped, "cooling down"):
            another.check()
        self.clock.sleep(121)
        another.check()

    def test_challenge_is_not_mistaken_for_expired_auth(self):
        self.reply = response(403, headers={"cf-mitigated": "challenge"})
        with self.assertRaisesRegex(policy.RequestStopped, "Cloudflare"):
            self.requests.get(self.sync.IR_DETAIL_URL + "/1")
        with self.assertRaises(policy.RequestStopped):
            self.cooldown.check()

    def test_success_status_cloudflare_challenge_is_blocked(self):
        self.reply = response(200, headers={"content-type": "text/html"},
            text='<title>Just a moment</title> Cloudflare /cdn-cgi/challenge-platform/')
        with self.assertRaisesRegex(policy.RequestStopped, "Cloudflare"):
            self.requests.get(self.sync.IR_DETAIL_URL + "/1")

    def test_http_failures_and_invalid_json_are_not_cached_or_retried(self):
        for reply in (response(401), response(403), response(503), response(302),
                      response(200, body={"Error": "private error"}), response(200, body="invalid")):
            with self.subTest(status=reply.status, body=reply.json.return_value):
                requests = policy.MeasuredRequests(self.client, self.sync, self.cooldown, self.cancel,
                    clock=self.clock, sleep=self.clock.sleep)
                self.reply = reply
                with self.assertRaises(policy.RequestStopped):
                    requests.post(self.sync.IRGRID_URL, params={"search": "test"}, data="[]")
                self.assertFalse(requests.cache)
                with self.assertRaises(policy.RequestStopped):
                    requests.get(self.sync.IR_DETAIL_URL + "/2")
                self.assertEqual(requests.requests, 1)

    def test_network_exception_does_not_leak_private_diagnostics(self):
        self.client.get.side_effect = RuntimeError("Authorization: Bearer private")
        with self.assertRaises(policy.RequestStopped) as caught:
            self.requests.get(self.sync.IR_DETAIL_URL + "/1")
        self.assertNotIn("private", str(caught.exception))

    def test_cancellation_interrupts_pacing_before_next_request(self):
        self.requests.get(self.sync.IR_DETAIL_URL + "/1")
        def cancel_after_first_wait():
            if self.clock() > 1000.25:
                raise worker.Cancelled()
        self.requests.check_cancel = cancel_after_first_wait
        with self.assertRaises(worker.Cancelled):
            self.requests.get(self.sync.IR_DETAIL_URL + "/2")
        self.assertEqual(len(self.starts), 1)

    def test_retry_after_seconds_dates_and_invalid_values(self):
        self.assertEqual(policy.retry_seconds("120"), 120)
        self.assertEqual(policy.retry_seconds("Thu, 01 Jan 1970 00:02:00 GMT", now=60), 60)
        for value in ("bad", "nan", "inf", "-5", None):
            self.assertEqual(policy.retry_seconds(value), 0)

    def test_reuse_preserves_distinct_vendor_selection(self):
        wb = Workbook()
        self.addCleanup(wb.close)
        ws = wb.active
        ws.append([core.H_VENDOR, core.H_INVOICE, core.H_STATUS, core.H_IR, core.H_PO])
        ws.append(["Alpha Plumbing", "INV-1", None, None, None])
        ws.append(["Beta Electrical", "INV-1", None, None, None])
        matches = [{"Invoice": "INV-1", "PayeeName": "Alpha Plumbing", "Inv_hmy": 101, "PO_hmy": 201},
                   {"Invoice": "INV-1", "PayeeName": "Beta Electrical", "Inv_hmy": 102, "PO_hmy": 202}]
        def send(url, **kwargs):
            self.starts.append(self.clock())
            self.clock.sleep(0.2)
            return response(body={"Records": matches}) if url == core.IRGRID_URL else response(body=[])
        self.client.post.side_effect = send
        requests = policy.MeasuredRequests(self.client, core, self.cooldown, self.cancel,
            clock=self.clock, sleep=self.clock.sleep)
        core.sync_sheet(ws, "Payments", SimpleNamespace(request=requests), {}, {}, {}, {}, {}, {}, {})
        self.assertEqual((ws.cell(2, 4).value, ws.cell(2, 5).value), (101, 201))
        self.assertEqual((ws.cell(3, 4).value, ws.cell(3, 5).value), (102, 202))
        self.assertEqual(sum(call.args[0] == core.IRGRID_URL for call in self.client.post.call_args_list), 1)
        self.assertEqual(requests.summary()["reused_searches"], 1)

    def test_date_retry_after_and_html_block_persist_wait(self):
        with patch.object(policy.time, "time", return_value=0):
            blocked = response(429, headers={"retry-after": "Thu, 01 Jan 1970 00:20:00 GMT"})
            with self.assertRaisesRegex(policy.RequestStopped, "1200 seconds"):
                policy.guard_response(blocked, self.cooldown)
        self.clock.sleep(301)
        with self.assertRaises(policy.RequestStopped):
            self.cooldown.check()
        denied = response(403, headers={"content-type": "text/html"}, text="Cloudflare Error 1020: Access denied")
        with self.assertRaisesRegex(policy.RequestStopped, "Cloudflare"):
            policy.guard_response(denied, self.cooldown, auth_probe=True)

    def test_session_probe_throttle_never_opens_browser(self):
        self.sync.AUTH_STATE_PATH = str(Path(self.temp.name) / "auth.json")
        self.sync.SESSION_DIR = str(Path(self.temp.name) / "browser")
        Path(self.sync.AUTH_STATE_PATH).write_text(json.dumps({"cookies": [{"name": "fixture"}]}))
        self.sync.get_auth_headers = lambda _: {"Authorization": "Bearer fixture"}
        playwright = MagicMock()
        playwright.request.new_context.return_value.post.return_value = response(429, headers={"retry-after": "600"})
        with self.assertRaisesRegex(policy.RequestStopped, "600 seconds"):
            with session_transport.open_session(playwright, self.sync, None, None, MagicMock(), lambda: None):
                self.fail("Blocked session must not reach workbook processing")
        playwright.chromium.launch_persistent_context.assert_not_called()
        playwright.request.new_context.return_value.post.assert_called_once()
        # Restarting cannot evade cooldown or even open a browser.
        playwright.reset_mock()
        with self.assertRaisesRegex(policy.RequestStopped, "cooling down"):
            with session_transport.open_session(playwright, self.sync, None, None, MagicMock(), lambda: None):
                pass
        playwright.request.new_context.assert_not_called()
        playwright.chromium.launch_persistent_context.assert_not_called()


class BlockedWorkbookTests(unittest.TestCase):
    def test_throttle_mid_row_leaves_workbook_and_backup_untouched(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            path = root / "registry.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append([core.H_VENDOR, core.H_INVOICE, core.H_STATUS, core.H_IR, core.H_PO, core.H_CH_NUM, core.H_CH_DATE])
            ws.append(["Vendor", "INV-1", "Pending", 101, 201, None, None])
            wb.save(path)
            wb.close()
            original = path.read_bytes()
            client = MagicMock()
            client.post.side_effect = [response(200, body=[{"Title": "Paid", "RevisedDate": "2026-10-01"}])]
            # EFT detail failure is inside a broad core catch: it must propagate,
            # not silently fall back to a cheque status and save it.
            client.get.return_value = response(429, headers={"retry-after": "120"})
            events = []
            clock = Clock()
            real_policy = policy.MeasuredRequests
            def paced(client, sync, cooldown, check_cancel):
                return real_policy(client, sync, cooldown, check_cancel, clock=clock, sleep=clock.sleep)
            worker.CANCEL.clear()
            with contextlib.ExitStack() as stack:
                for item in (
                    patch.object(worker, "emit", side_effect=lambda kind, **data: events.append({"type": kind, **data})),
                    patch.object(core, "APP_DATA_DIR", str(root / "session")),
                    patch.object(core, "SESSION_DIR", str(root / "session" / "browser")),
                    patch.object(core, "AUTH_STATE_PATH", str(root / "session" / "auth.json")),
                    patch("session_transport.open_session", return_value=contextlib.nullcontext((SimpleNamespace(request=client), {}))),
                    patch("playwright.sync_api.sync_playwright", return_value=MagicMock()),
                    patch.object(core, "preload_irgrid", return_value={}),
                    patch.object(core, "preload_pogrid", return_value={}),
                    patch.object(worker, "MeasuredRequests", side_effect=paced)):
                    stack.enter_context(item)
                with self.assertRaises(policy.RequestStopped):
                    worker.run(path)
            self.assertEqual(path.read_bytes(), original)
            self.assertFalse((root / "Registry Backups").exists())
            self.assertFalse(any(event["type"] in ("complete", "row") for event in events))
            self.assertEqual(client.post.call_count, 1)
            self.assertEqual(client.get.call_count, 1)
            self.assertTrue(any("Sync API timing" in event.get("message", "") for event in events))


if __name__ == "__main__":
    unittest.main()
