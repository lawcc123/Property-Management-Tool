"""Sequential, measured read-only sync transport. No probing or automatic retries."""
from collections import defaultdict
from email.utils import parsedate_to_datetime
import hashlib
import json
import math
from pathlib import Path
import tempfile
import os
import time
from urllib.parse import urlsplit


class RequestStopped(RuntimeError):
    """Abort the entire sync; never turn throttling into a row-level result."""


class Cooldown:
    def __init__(self, path, clock=time.time):
        self.path = Path(path)
        self.clock = clock

    @classmethod
    def for_sync(cls, sync):
        origin = urlsplit(sync.IRGRID_URL).netloc.lower()
        scope = hashlib.sha256(origin.encode()).hexdigest()[:16]
        return cls(Path(sync.AUTH_STATE_PATH).parent / f"sync-cooldown-{scope}.json")

    def check(self):
        try:
            remaining = float(json.loads(self.path.read_text())["until"]) - self.clock()
        except FileNotFoundError:
            return
        except (OSError, ValueError, TypeError, KeyError):
            raise RequestStopped("The Yardi cooldown record could not be read. Check the local session folder before retrying; no workbook changes were saved.") from None
        if not math.isfinite(remaining):
            raise RequestStopped("The Yardi cooldown record is invalid. Check the local session folder before retrying.")
        if remaining > 0:
            raise RequestStopped(f"Yardi access is cooling down. Wait at least {math.ceil(remaining)} seconds before starting again. No workbook changes were saved.")

    def hold(self, seconds):
        staged = None
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            until = self.clock() + seconds
            try:
                previous = float(json.loads(self.path.read_text())["until"])
                if math.isfinite(previous):
                    until = max(until, previous)
            except (OSError, ValueError, TypeError, KeyError):
                pass
            descriptor, staged = tempfile.mkstemp(prefix="cooldown-", dir=self.path.parent)
            with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
                json.dump({"until": until}, stream)
            os.replace(staged, self.path)
            return True
        except OSError:
            return False
        finally:
            if staged and os.path.exists(staged):
                os.unlink(staged)


def retry_seconds(value, now=None):
    """Retry-After permits either seconds or an HTTP date. Ignore invalid values."""
    if not isinstance(value, str):
        return 0
    try:
        seconds = float(value)
    except ValueError:
        try:
            seconds = parsedate_to_datetime(value).timestamp() - (time.time() if now is None else now)
        except (ValueError, TypeError, OverflowError):
            return 0
    return max(0, seconds) if math.isfinite(seconds) else 0


def guard_response(response, cooldown, auth_probe=False):
    """Stop on blocks, never retry them or reopen login to work around them."""
    headers = {key.lower(): value for key, value in response.headers.items()}
    status = response.status
    content_type = headers.get("content-type", "").lower()
    text = ""
    if "html" in content_type or status in (403, 429):
        try:
            raw = response.text()
            text = raw[:131072].lower() if isinstance(raw, str) else ""
        except Exception:
            pass
    challenge = (headers.get("cf-mitigated", "").lower() == "challenge"
                 or "/cdn-cgi/challenge-platform/" in text or "cf-chl-" in text
                 or ("cloudflare" in text and ("just a moment" in text or "error 1015" in text
                       or "error 1020" in text or "attention required!" in text
                       or "sorry, you have been blocked" in text
                       or "you are being rate limited" in text)))
    if status == 429 or challenge:
        seconds = max(300 if challenge else 60, retry_seconds(headers.get("retry-after", "")))
        remembered = (cooldown() if callable(cooldown) else cooldown).hold(seconds)
        reason = "Cloudflare requires verification or has blocked requests" if challenge else "Yardi returned HTTP 429 (too many requests)"
        suffix = "" if remembered else " The cooldown could not be remembered; do not restart before this wait ends."
        raise RequestStopped(f"{reason}. Sync stopped without retrying. Wait at least {math.ceil(seconds)} seconds before trying again; no workbook changes were saved.{suffix}")
    if not auth_probe:
        if status in (401, 403):
            raise RequestStopped(f"Yardi returned HTTP {status}. Sync stopped; sign-in or access may need attention. No workbook changes were saved.")
        if not 200 <= status < 300 or "html" in content_type:
            # A server-requested pause must survive restarting the app too.
            seconds = retry_seconds(headers.get("retry-after", ""))
            if seconds:
                (cooldown() if callable(cooldown) else cooldown).hold(seconds)
            raise RequestStopped(f"Yardi returned an unexpected response (HTTP {status}). Sync stopped without retrying; no workbook changes were saved.")


class MeasuredRequests:
    """One caller, one request at a time, >=0.3s quiet time after each response.

    This retains the existing lookup pause and also covers previously unpaced
    detail calls. Only identical read-only targeted grid searches are cached.
    All caches die with the run; vendor/amount/date selection still runs per row.
    """
    uses_global_pacing = True

    def __init__(self, client, sync, cooldown, check_cancel, clock=time.monotonic, sleep=time.sleep):
        self.client, self.cooldown, self.check_cancel = client, cooldown, check_cancel
        self.clock, self.sleep = clock, sleep
        self.interval = 0.3
        self.next_request = 0
        self.stopped = None
        self.labels = {getattr(sync, key): label for key, label in (
            ("IRGRID_URL", "IR search"), ("POGRID_URL", "PO search"),
            ("IR_MILESTONE_URL", "IR milestones"), ("PO_MILESTONE_URL", "PO milestones"),
            ("IR_DETAIL_URL", "IR details"))}
        self.search_urls = {sync.IRGRID_URL, sync.POGRID_URL}
        self.cache = {}
        self.samples = defaultdict(list)
        self.cache_hits = self.wait_seconds = self.requests = 0

    def check(self):
        self.check_cancel()
        if self.stopped:
            raise self.stopped
        self.cooldown.check()

    def _call(self, method, url, **kwargs):
        self.check()
        # Exact request parameters/body, scoped to this authenticated client;
        # never persist URLs, search text, headers, cookies or response content.
        params = kwargs.get("params")
        cache_key = None
        if method == "post" and url in self.search_urls and isinstance(params, dict) and params.get("search"):
            cache_key = (url, json.dumps(params, sort_keys=True), kwargs.get("data"))
            if cache_key in self.cache:
                self.cache_hits += 1
                return self.cache[cache_key]
        waiting = self.clock()
        while self.clock() < self.next_request:
            self.check()
            self.sleep(max(0, min(0.1, self.next_request - self.clock())))
        self.wait_seconds += self.clock() - waiting
        self.check()
        start = self.clock()
        self.requests += 1
        label = next((value for key, value in self.labels.items() if url == key or url.startswith(key + "/")), "Other read")
        response = None
        try:
            response = getattr(self.client, method)(url, **kwargs)
            guard_response(response, self.cooldown)
            # Validate before caching. Malformed responses must never become
            # a negative match or trigger further fallback requests.
            body = response.json()
            if not isinstance(body, (dict, list)) or (isinstance(body, dict) and any(
                    body.get(key) for key in ("error", "Error", "Exception"))):
                raise RequestStopped("Yardi returned an unexpected API result. Sync stopped; no workbook changes were saved.")
            if cache_key is not None:
                self.cache[cache_key] = response
            return response
        except RequestStopped as exc:
            self.stopped = exc
            if response is not None:
                response.dispose()
            raise
        except Exception:
            self.stopped = RequestStopped("A Yardi request failed or returned unreadable data. Sync stopped without retrying; no workbook changes were saved.")
            if response is not None:
                response.dispose()
            raise self.stopped from None
        finally:
            end = self.clock()
            self.samples[label].append(end - start)
            self.next_request = end + self.interval

    def get(self, url, **kwargs):
        return self._call("get", url, **kwargs)

    def post(self, url, **kwargs):
        return self._call("post", url, **kwargs)

    def summary(self):
        endpoints = {}
        for name, values in sorted(self.samples.items()):
            ordered = sorted(values)
            endpoints[name] = {"requests": len(values), "seconds": round(sum(values), 3),
                               "p50_ms": round(ordered[(len(ordered)-1)//2]*1000),
                               "p95_ms": round(ordered[math.ceil(len(ordered)*0.95)-1]*1000)}
        return {"requests": self.requests, "reused_searches": self.cache_hits,
                "network_seconds": round(sum(sum(values) for values in self.samples.values()), 3),
                "pacing_seconds": round(self.wait_seconds, 3), "endpoints": endpoints}
