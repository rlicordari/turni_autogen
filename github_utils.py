# -*- coding: utf-8 -*-
"""GitHub helpers (repo storage via the Contents API).

All app traffic shares ONE token, so every call goes through a small gateway:

- rate limits (403/429 with rate-limit markers), 5xx and network errors are
  retried, honouring `retry-after` / `x-ratelimit-reset`, within `max_wait`;
- when GitHub asks to wait longer than `max_wait`, GithubUnavailable is raised
  with the time to retry at, and later calls fail fast until then (no hammering);
- writes are serialized process-wide and paced below GitHub's limit of 80
  content-creating requests per minute: all sessions of the Streamlit app live in
  this process, so our own commits never race on the branch.

Conflicts (409/422 sha mismatch) and permission errors are NOT retried here:
callers must re-read and decide.
"""

from __future__ import annotations

import base64
import json
import threading
import time
from dataclasses import dataclass
from typing import Optional

import requests

WRITE_MIN_INTERVAL = 1.0   # seconds between commits: max 60/min (GitHub limit: 80/min)
DEFAULT_MAX_WAIT = 25.0    # interactive calls: how long we are willing to wait
_BACKOFF_BASE = 1.0

# Test hooks
_http = requests.request
_sleep = time.sleep
_now = time.time

_write_lock = threading.Lock()
_state_lock = threading.Lock()
_state = {"blocked_until": 0.0, "last_write": 0.0}


class GithubUnavailable(RuntimeError):
    """GitHub is rate limiting or unreachable: retry at `retry_at` (epoch seconds)."""

    def __init__(self, message: str, retry_at: float):
        super().__init__(message)
        self.retry_at = retry_at


@dataclass
class GithubFile:
    text: str
    sha: str


def reset_state() -> None:
    with _state_lock:
        _state["blocked_until"] = 0.0
        _state["last_write"] = 0.0


def blocked_until() -> float:
    with _state_lock:
        return _state["blocked_until"]


def _block(until: float) -> None:
    with _state_lock:
        _state["blocked_until"] = max(_state["blocked_until"], until)


def _is_rate_limited(r: requests.Response) -> bool:
    if r.status_code == 429:
        return True
    if r.status_code != 403:
        return False
    if r.headers.get("x-ratelimit-remaining") == "0" or r.headers.get("retry-after"):
        return True
    return "rate limit" in (r.text or "").lower()


def _rate_limit_wait(r: requests.Response, attempt: int) -> float:
    retry_after = r.headers.get("retry-after")
    if retry_after:
        try:
            return max(1.0, float(retry_after))
        except ValueError:
            pass
    if r.headers.get("x-ratelimit-remaining") == "0" and r.headers.get("x-ratelimit-reset"):
        try:
            return max(1.0, float(r.headers["x-ratelimit-reset"]) - _now() + 1)
        except ValueError:
            pass
    # GitHub: "wait at least one minute", exponentially more if it persists
    return 60.0 * (2 ** attempt)


def request(method: str, url: str, *, max_wait: Optional[float] = None, write: bool = False, **kwargs) -> requests.Response:
    """HTTP call with retries on rate limits / 5xx / network errors (see module doc)."""
    max_wait = DEFAULT_MAX_WAIT if max_wait is None else max_wait
    deadline = _now() + max_wait
    attempt = 0
    while True:
        now = _now()
        until = blocked_until()
        if until > now:
            if until > deadline:
                raise GithubUnavailable(
                    f"GitHub limita le richieste: riprova dopo {int(until - now)} secondi.", until
                )
            _sleep(until - now)
        if write:
            with _write_lock:
                gap = _state["last_write"] + WRITE_MIN_INTERVAL - _now()
                if gap > 0:
                    _sleep(gap)
                try:
                    r = _http(method, url, **kwargs)
                    err = None
                except (requests.ConnectionError, requests.Timeout) as e:
                    r, err = None, e
                with _state_lock:
                    _state["last_write"] = _now()
        else:
            try:
                r = _http(method, url, **kwargs)
                err = None
            except (requests.ConnectionError, requests.Timeout) as e:
                r, err = None, e

        if r is not None and _is_rate_limited(r):
            wait = _rate_limit_wait(r, attempt)
            _block(_now() + wait)
            if _now() + wait > deadline:
                raise GithubUnavailable(
                    f"GitHub limita le richieste (HTTP {r.status_code}): riprova tra {int(wait)} secondi.",
                    _now() + wait,
                )
        elif r is not None and r.status_code < 500:
            return r
        else:
            wait = _BACKOFF_BASE * (2 ** attempt)
            if _now() + wait > deadline:
                what = f"HTTP {r.status_code}" if r is not None else f"{type(err).__name__}: {err}"
                raise GithubUnavailable(f"GitHub non raggiungibile ({what}).", _now() + wait)
            _sleep(wait)
        attempt += 1


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json",
    }


def _raise_for_status(r: requests.Response) -> None:
    if r.status_code == 404:
        raise requests.HTTPError(
            f"404 Not Found for {r.url}. Check owner/repo/path/branch and token permissions (Contents: read/write).",
            response=r,
        )
    r.raise_for_status()


def get_file(
    owner: str,
    repo: str,
    path: str,
    token: str,
    branch: str = "main",
    timeout_s: int = 20,
    max_wait: Optional[float] = None,
) -> Optional[GithubFile]:
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    r = request("GET", url, headers=_headers(token), params={"ref": branch}, timeout=timeout_s, max_wait=max_wait)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    data = r.json()
    content_b64 = (data.get("content", "") or "").replace("\n", "")
    raw = base64.b64decode(content_b64.encode("utf-8")) if content_b64 else b""
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("utf-8", errors="replace")
    return GithubFile(text=text, sha=data.get("sha", ""))


def list_dir(
    owner: str,
    repo: str,
    path: str,
    token: str,
    branch: str = "main",
    timeout_s: int = 20,
    max_wait: Optional[float] = None,
) -> list[dict]:
    """List files in a GitHub directory. Returns [] if path does not exist or is a plain file."""
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    r = request("GET", url, headers=_headers(token), params={"ref": branch}, timeout=timeout_s, max_wait=max_wait)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    data = r.json()
    return data if isinstance(data, list) else []


def put_file(
    owner: str,
    repo: str,
    path: str,
    token: str,
    message: str,
    text: str,
    branch: str = "main",
    sha: Optional[str] = None,
    timeout_s: int = 20,
    max_wait: Optional[float] = None,
) -> dict:
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    payload = {
        "message": message,
        "content": base64.b64encode(text.encode("utf-8")).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha
    r = request("PUT", url, headers=_headers(token), json=payload, timeout=timeout_s, max_wait=max_wait, write=True)
    _raise_for_status(r)
    try:
        return r.json()
    except json.JSONDecodeError:
        return {}
