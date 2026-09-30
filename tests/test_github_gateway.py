import threading
import time
import unittest
from unittest import mock

import requests

import github_utils


class FakeClock:
    def __init__(self):
        self.t = 1_000_000.0
        self.lock = threading.Lock()

    def now(self):
        with self.lock:
            return self.t

    def sleep(self, s):
        with self.lock:
            self.t += max(0.0, s)


def response(status, body=b'{"content": "", "sha": "abc"}', headers=None):
    r = requests.models.Response()
    r.status_code = status
    r._content = body
    r.headers.update(headers or {})
    r.url = "https://api.github.com/repos/o/r/contents/x"
    return r


class GatewayTests(unittest.TestCase):
    def setUp(self):
        self.clock = FakeClock()
        self.calls = []
        self.replies = []
        github_utils.reset_state()
        for target, value in (
            ("_now", self.clock.now),
            ("_sleep", self.clock.sleep),
            ("_http", self.fake_http),
        ):
            p = mock.patch.object(github_utils, target, value)
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(github_utils.reset_state)

    def fake_http(self, method, url, **kw):
        self.calls.append((method, self.clock.now()))
        reply = self.replies.pop(0) if self.replies else response(200)
        if isinstance(reply, Exception):
            raise reply
        return reply

    def test_secondary_rate_limit_waits_retry_after_then_succeeds(self):
        self.replies = [
            response(403, b'{"message": "You have exceeded a secondary rate limit"}', {"retry-after": "7"}),
            response(200),
        ]

        gf = github_utils.get_file("o", "r", "x", "t", max_wait=30)

        self.assertEqual(gf.sha, "abc")
        self.assertEqual(len(self.calls), 2)
        self.assertGreaterEqual(self.calls[1][1] - self.calls[0][1], 7)

    def test_primary_limit_uses_reset_header(self):
        reset = int(self.clock.now()) + 12
        self.replies = [
            response(403, b'{"message": "API rate limit exceeded"}',
                     {"x-ratelimit-remaining": "0", "x-ratelimit-reset": str(reset)}),
            response(200),
        ]

        github_utils.get_file("o", "r", "x", "t", max_wait=30)

        self.assertGreaterEqual(self.calls[1][1], reset)

    def test_429_server_errors_and_network_errors_are_retried(self):
        self.replies = [
            response(429, b"{}", {"retry-after": "1"}),
            response(502, b"bad gateway"),
            requests.ConnectionError("reset by peer"),
            requests.Timeout("slow"),
            response(200),
        ]

        gf = github_utils.get_file("o", "r", "x", "t", max_wait=60)

        self.assertEqual(gf.sha, "abc")
        self.assertEqual(len(self.calls), 5)

    def test_gives_up_with_github_unavailable_when_wait_exceeds_budget(self):
        self.replies = [response(403, b'{"message": "secondary rate limit"}', {"retry-after": "60"})]

        with self.assertRaises(github_utils.GithubUnavailable) as ctx:
            github_utils.put_file("o", "r", "x", "t", "msg", "text", max_wait=10)

        self.assertGreaterEqual(ctx.exception.retry_at, self.clock.now() + 50)
        self.assertEqual(len(self.calls), 1)

    def test_known_block_fails_fast_without_calling_github(self):
        self.replies = [response(429, b"{}", {"retry-after": "120"})]
        with self.assertRaises(github_utils.GithubUnavailable):
            github_utils.get_file("o", "r", "x", "t", max_wait=5)

        with self.assertRaises(github_utils.GithubUnavailable):
            github_utils.put_file("o", "r", "x", "t", "msg", "text", max_wait=5)

        self.assertEqual(len(self.calls), 1)

    def test_plain_403_permission_error_is_not_retried(self):
        self.replies = [response(403, b'{"message": "Resource not accessible by personal access token"}')]

        with self.assertRaises(requests.HTTPError):
            github_utils.get_file("o", "r", "x", "t", max_wait=30)

        self.assertEqual(len(self.calls), 1)

    def test_conflict_is_returned_to_caller_not_retried_blindly(self):
        self.replies = [response(409, b'{"message": "sha does not match"}')]

        with self.assertRaises(requests.HTTPError) as ctx:
            github_utils.put_file("o", "r", "x", "t", "msg", "text", sha="old", max_wait=30)

        self.assertEqual(ctx.exception.response.status_code, 409)
        self.assertEqual(len(self.calls), 1)

    def test_writes_are_paced_under_github_content_limit(self):
        for i in range(5):
            github_utils.put_file("o", "r", f"f{i}", "t", "msg", "text", max_wait=30)

        times = [t for _m, t in self.calls]
        gaps = [b - a for a, b in zip(times, times[1:])]
        self.assertTrue(all(g >= github_utils.WRITE_MIN_INTERVAL for g in gaps), gaps)
        self.assertLess(60 / github_utils.WRITE_MIN_INTERVAL, 80)


class GatewayConcurrencyTests(unittest.TestCase):
    """Real threads and real time: writes from many sessions never overlap."""

    def setUp(self):
        github_utils.reset_state()
        self.addCleanup(github_utils.reset_state)
        self.in_flight = 0
        self.max_in_flight = 0
        self.lock = threading.Lock()

    def slow_http(self, method, url, **kw):
        if method == "PUT":
            with self.lock:
                self.in_flight += 1
                self.max_in_flight = max(self.max_in_flight, self.in_flight)
            time.sleep(0.01)
            with self.lock:
                self.in_flight -= 1
        return response(201, b'{"content": {"sha": "s"}, "commit": {"sha": "c"}}')

    def test_concurrent_writes_are_serialized(self):
        with mock.patch.object(github_utils, "_http", self.slow_http), \
                mock.patch.object(github_utils, "WRITE_MIN_INTERVAL", 0.0):
            threads = [
                threading.Thread(target=github_utils.put_file, args=("o", "r", f"f{i}", "t", "m", "x"))
                for i in range(20)
            ]
            for t in threads:
                t.start()
            for t in threads:
                t.join()

        self.assertEqual(self.max_in_flight, 1)


if __name__ == "__main__":
    unittest.main()
