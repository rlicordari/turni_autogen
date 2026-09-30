"""Run the real app locally on a COPY of the data (no GitHub writes, no emails).

    .venv/bin/streamlit run scripts/run_local_sandbox.py

Data copy: .local_sandbox/repo (git clone --depth 1 of the data repo).
SMTP is replaced too: mails are only printed to the console.
"""

import os
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
for p in (str(ROOT), str(ROOT / "scripts")):
    if p not in sys.path:
        sys.path.insert(0, p)

import smtplib  # noqa: E402

import github_utils  # noqa: E402
from sandbox_github import sandbox  # noqa: E402


class _SandboxSMTP:
    """No mail ever leaves the sandbox, even with credentials saved in the admin panel."""

    def __init__(self, host, port, timeout=20):
        self.host = host

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def ehlo(self):
        pass

    def starttls(self):
        pass

    def login(self, user, password):
        print(f"[SANDBOX] login SMTP simulato ({user} @ {self.host})", flush=True)

    def send_message(self, msg):
        print(f"[SANDBOX] mail NON inviata: a={msg['To']} cc={msg['Cc']} oggetto={msg['Subject']}", flush=True)

    def close(self):
        pass


os.environ["TURNI_SANDBOX"] = "1"
smtplib.SMTP = _SandboxSMTP
os.environ.setdefault("TURNI_QUEUE_PERSIST_PATH", str(ROOT / ".local_sandbox" / "save_queue.json"))
github_utils._http = sandbox()
runpy.run_path(str(ROOT / "streamlit_app.py"), run_name="__main__")
