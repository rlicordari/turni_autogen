"""Run the real app locally on a COPY of the data (no GitHub writes, no emails).

    .venv/bin/streamlit run scripts/run_local_sandbox.py

Data copy: .local_sandbox/repo (git clone --depth 1 of the data repo).
"""

import os
import runpy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
for p in (str(ROOT), str(ROOT / "scripts")):
    if p not in sys.path:
        sys.path.insert(0, p)

import github_utils  # noqa: E402
from sandbox_github import sandbox  # noqa: E402

os.environ["TURNI_SANDBOX"] = "1"
os.environ.setdefault("TURNI_QUEUE_PERSIST_PATH", str(ROOT / ".local_sandbox" / "save_queue.json"))
github_utils._http = sandbox()
runpy.run_path(str(ROOT / "streamlit_app.py"), run_name="__main__")
