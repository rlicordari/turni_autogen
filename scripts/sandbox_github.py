"""GitHub Contents API served from a local folder (sandbox for local runs).

The app keeps calling github_utils as usual; only the HTTP layer is replaced, so
reads come from the local copy of the data repo and writes stay on disk.
Nothing reaches GitHub, nothing is emailed (no SMTP in local secrets).
"""

from __future__ import annotations

import base64
import hashlib
import json
import threading
import uuid
from pathlib import Path
from urllib.parse import unquote

import requests

ROOT = Path(__file__).resolve().parents[1]
REPO_DIR = ROOT / ".local_sandbox" / "repo"


def _response(status: int, payload, url: str) -> requests.Response:
    r = requests.models.Response()
    r.status_code = status
    r._content = json.dumps(payload).encode()
    r.url = url
    return r


def _blob_sha(data: bytes) -> str:
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


class LocalFolderGithub:
    def __init__(self, root: Path):
        self.root = root.resolve()
        self.lock = threading.Lock()

    def _target(self, path: str) -> Path:
        target = (self.root / path).resolve()
        if self.root not in target.parents and target != self.root:
            raise ValueError(f"path outside sandbox: {path}")
        return target

    def __call__(self, method: str, url: str, **kw) -> requests.Response:
        path = unquote(url.split("/contents/", 1)[1]).strip("/")
        target = self._target(path)
        with self.lock:
            if method == "GET":
                if target.is_file():
                    data = target.read_bytes()
                    return _response(200, {"content": base64.b64encode(data).decode(), "sha": _blob_sha(data)}, url)
                if target.is_dir():
                    items = [
                        {"name": p.name, "path": f"{path}/{p.name}", "type": "file"}
                        for p in sorted(target.iterdir())
                        if p.is_file() and not p.name.startswith(".")
                    ]
                    return _response(200, items, url)
                return _response(404, {"message": "Not Found"}, url)
            if method == "PUT":
                payload = kw.get("json") or {}
                sha = payload.get("sha")
                current = _blob_sha(target.read_bytes()) if target.is_file() else None
                if current is not None and not sha:
                    return _response(422, {"message": 'Invalid request. "sha" wasn\'t supplied.'}, url)
                if (current is None and sha) or (current is not None and sha != current):
                    return _response(409, {"message": f"{path} does not match {sha}"}, url)
                data = base64.b64decode(payload["content"])
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
                print(f"[SANDBOX] scritto {path} — {payload.get('message', '')}", flush=True)
                return _response(201, {"content": {"sha": _blob_sha(data)}, "commit": {"sha": uuid.uuid4().hex}}, url)
        return _response(405, {"message": "method not allowed"}, url)


_SANDBOX = None


def sandbox() -> LocalFolderGithub:
    global _SANDBOX
    if _SANDBOX is None:
        if not REPO_DIR.is_dir():
            raise SystemExit(
                "Copia dati mancante: git clone --depth 1 https://github.com/rlicordari/Turni-ind .local_sandbox/repo"
            )
        _SANDBOX = LocalFolderGithub(REPO_DIR)
    return _SANDBOX
