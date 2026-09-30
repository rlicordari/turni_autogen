"""Test doubles: GitHub Contents API at HTTP level (with chaos), SMTP, clock."""

from __future__ import annotations

import base64
import json
import random
import smtplib
import threading
from urllib.parse import unquote

import requests

API = "https://api.github.com/repos/"


def http_response(status: int, payload=None, headers=None, url: str = "") -> requests.Response:
    r = requests.models.Response()
    r.status_code = status
    r._content = json.dumps(payload if payload is not None else {}).encode()
    r.headers.update(headers or {})
    r.url = url
    return r


class FakeClock:
    """Shared fake time: sleep() advances it instantly (retries never wait for real)."""

    def __init__(self, start: float = 1_800_000_000.0):
        self.t = start
        self.lock = threading.Lock()

    def now(self) -> float:
        with self.lock:
            return self.t

    def sleep(self, seconds: float) -> None:
        with self.lock:
            self.t += max(0.0, seconds)


class FakeGithubHTTP:
    """Minimal GitHub Contents API: GET file/dir, PUT with optimistic lock on sha.

    chaos: probability that a request fails before reaching the "server" with a
    rate limit (403/429 + retry-after), a 502, a timeout or a connection reset.
    down: every request is answered with a rate limit (GitHub saturated).
    """

    def __init__(self, chaos: float = 0.0, seed: int = 7):
        self.files: dict[str, tuple[str, str]] = {}
        self.n = 0
        self.lock = threading.Lock()
        self.rng = random.Random(seed)
        self.chaos = chaos
        self.down = False
        self.put_in_flight = 0
        self.max_put_in_flight = 0
        self.requests = 0
        self.commits: list[tuple[str, str]] = []  # (path, message)

    # -- helpers for tests --
    def seed(self, path: str, text: str) -> None:
        with self.lock:
            self.n += 1
            self.files[path] = (text, f"sha{self.n}")

    def text(self, path: str) -> str:
        return self.files.get(path, ("", ""))[0]

    # -- HTTP --
    def __call__(self, method: str, url: str, **kw) -> requests.Response:
        with self.lock:
            self.requests += 1
            roll = self.rng.random()
            pick = self.rng.random()
        if self.down:
            return http_response(403, {"message": "You have exceeded a secondary rate limit"}, {"retry-after": "60"}, url)
        if roll < self.chaos:
            if pick < 0.3:
                return http_response(403, {"message": "You have exceeded a secondary rate limit"}, {"retry-after": "2"}, url)
            if pick < 0.5:
                return http_response(429, {"message": "Too Many Requests"}, {"retry-after": "1"}, url)
            if pick < 0.7:
                return http_response(502, {"message": "Bad Gateway"}, url=url)
            if pick < 0.85:
                raise requests.Timeout("read timed out")
            raise requests.ConnectionError("connection reset by peer")

        path = unquote(url.split("/contents/", 1)[1])
        if method == "GET":
            return self._get(path, url)
        if method == "PUT":
            with self.lock:
                self.put_in_flight += 1
                self.max_put_in_flight = max(self.max_put_in_flight, self.put_in_flight)
                lose_response = self.rng.random() < self.chaos * 0.3
            try:
                reply = self._put(path, kw.get("json") or {}, url)
            finally:
                with self.lock:
                    self.put_in_flight -= 1
            if lose_response and reply.status_code < 300:
                # Commit applicato da GitHub, ma la risposta non arriva al client.
                raise requests.Timeout("response lost after commit")
            return reply
        return http_response(405, {"message": "method"}, url=url)

    def _get(self, path: str, url: str) -> requests.Response:
        with self.lock:
            if path in self.files:
                text, sha = self.files[path]
                return http_response(200, {"content": base64.b64encode(text.encode()).decode(), "sha": sha}, url=url)
            prefix = path.rstrip("/") + "/"
            children = [
                {"name": p[len(prefix):], "path": p, "type": "file"}
                for p in sorted(self.files)
                if p.startswith(prefix) and "/" not in p[len(prefix):]
            ]
            if children:
                return http_response(200, children, url=url)
        return http_response(404, {"message": "Not Found"}, url=url)

    def _put(self, path: str, payload: dict, url: str) -> requests.Response:
        with self.lock:
            current = self.files.get(path)
            sha = payload.get("sha")
            if current is not None and not sha:
                return http_response(422, {"message": 'Invalid request. "sha" wasn\'t supplied.'}, url=url)
            if current is not None and sha != current[1]:
                return http_response(409, {"message": f"{path} does not match {sha}"}, url=url)
            if current is None and sha:
                return http_response(409, {"message": f"{path} does not match {sha}"}, url=url)
            self.n += 1
            text = base64.b64decode(payload["content"]).decode()
            self.files[path] = (text, f"sha{self.n}")
            self.commits.append((path, payload.get("message", "")))
            return http_response(201, {"content": {"sha": f"sha{self.n}"}, "commit": {"sha": f"commit{self.n:07d}"}}, url=url)


class FakeSMTP:
    sent: list = []
    logins: list = []          # (username, password) of every login
    reject_login = False       # True: the server refuses the credentials
    lock = threading.Lock()

    def __init__(self, host, port, timeout=20):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def ehlo(self):
        pass

    def starttls(self):
        pass

    def login(self, user, password):
        with FakeSMTP.lock:
            FakeSMTP.logins.append((user, password))
        if FakeSMTP.reject_login:
            raise smtplib.SMTPAuthenticationError(535, b"5.7.8 Username and Password not accepted")

    def send_message(self, msg):
        with FakeSMTP.lock:
            FakeSMTP.sent.append(msg)
