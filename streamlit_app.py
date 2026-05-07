import io
import tempfile
import time
import traceback
import csv
import json
import requests
import random
import re
import uuid
import os
import base64
import hashlib
import hmac
import smtplib
from email.message import EmailMessage
from datetime import date, datetime, timezone, timedelta
from pathlib import Path
from collections.abc import Mapping

import streamlit as st
import pandas as pd
import plotly.express as px
import yaml

# Local modules
import github_utils
import unavailability_store as ustore
import xlsx_utils
import shift_history as sh

# Import generator
import turni_generator as tg

APP_BUILD = "2026-02-01-ui-v7"

# ---- Concurrency & session safety (doctor mode) ----
# We implement a per-doctor "lease" file on GitHub:
#   - new login overwrites the lease (kicking out previous sessions)
#   - older sessions detect the mismatch and are forced to log out
# Additionally, saves are optimistic-concurrency safe via GitHub SHA + retries
# and are verified by a read-back signature check.

DOCTOR_SESSION_TTL_MINUTES = 20
DOCTOR_SESSION_CHECK_SECONDS = 5          # throttle for lease mismatch checks
DOCTOR_SESSION_HEARTBEAT_SECONDS = 60     # throttle for lease keep-alive writes

def _utc_now_iso() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"

# ---- UI flash messages (persist across reruns) ----
def _unav_flash_key(doctor: str) -> str:
    return f"unav_flash__{doctor}"

def set_unav_flash(doctor: str, kind: str, msg: str, details: str | None = None) -> None:
    """Persist a message (success/error/info) so it doesn't disappear on rerun."""
    st.session_state[_unav_flash_key(doctor)] = {
        "kind": kind,
        "msg": msg,
        "details": details,
        "ts": _utc_now_iso(),
    }

def render_unav_flash(doctor: str) -> None:
    key = _unav_flash_key(doctor)
    f = st.session_state.get(key)
    if not isinstance(f, dict):
        return

    cols = st.columns([12, 1])
    with cols[0]:
        kind = str(f.get("kind") or "info")
        msg = str(f.get("msg") or "")
        if kind == "success":
            st.success(msg)
        elif kind == "error":
            st.error(msg)
        else:
            st.info(msg)

        details = f.get("details")
        if details:
            with st.expander("Dettagli"):
                st.code(str(details))

    with cols[1]:
        if st.button("✖", key=f"dismiss__{key}"):
            st.session_state.pop(key, None)
            st.rerun()

# ---- Indisponibilità: fasce ammesse e normalizzazione (per compatibilità con valori "storici") ----
FASCIA_OPTIONS = ["Mattina", "Pomeriggio", "Notte", "Diurno", "Tutto il giorno", "Ferie"]
AVAIL_FASCIA_OPTIONS = [f for f in FASCIA_OPTIONS if f != "Ferie"]
MAX_WEEKEND_DAYS = 2  # max sabati e max domeniche per mese (separatamente)

def normalize_fascia(val: object) -> tuple[str, bool, bool]:
    """Return (canonical_value, changed, unknown).

    - changed: value was recognized but normalized (e.g., 'matt' -> 'Mattina')
    - unknown: value wasn't recognized; we default to 'Tutto il giorno' but we warn the user.
    """
    if val is None:
        return "", False, False
    s = str(val).strip()
    if not s:
        return "", False, False
    key = s.casefold().strip()
    key = " ".join(key.split())  # collapse whitespace

    # direct matches (case-insensitive)
    direct = {
        "mattina": "Mattina",
        "pomeriggio": "Pomeriggio",
        "notte": "Notte",
        "diurno": "Diurno",
        "tutto il giorno": "Tutto il giorno",
        "tutto giorno": "Tutto il giorno",
        "all day": "Tutto il giorno",
        "giornata intera": "Tutto il giorno",
        "ferie": "Ferie",
    }
    if key in direct:
        canon = direct[key]
        return canon, canon != s, False

    # fuzzy matches
    if "tutto" in key or "all" in key or "intera" in key:
        return "Tutto il giorno", True, False
    if "diurn" in key or "daytime" in key or key == "d":
        return "Diurno", True, False
    if "matt" in key or "morning" in key or key in {"am", "a.m."}:
        return "Mattina", True, False
    if "pome" in key or "pom" in key or "afternoon" in key or key in {"pm", "p.m."}:
        return "Pomeriggio", True, False
    if "nott" in key or "night" in key or key == "n":
        return "Notte", True, False
    if "ferie" in key or "vacan" in key or "holiday" in key or "leave" in key:
        return "Ferie", True, False

    # unknown
    return "Tutto il giorno", True, True
# ---------------- Page config & style ----------------
st.set_page_config(
    page_title="UOC Cardiologia con UTIC - Turni",
    page_icon="🗓️",
    layout="wide",
)

st.markdown(
    """
<style>
/* Tidy up spacing */
.block-container { padding-top: 1.2rem; padding-bottom: 2.5rem; }
h1 { margin-bottom: 0.2rem; }
 .small-muted { opacity: 0.75; font-size: 0.92rem; }
.kpi { padding: 0.75rem 0.9rem; border-radius: 0.75rem; border: 1px solid rgba(128, 128, 128, 0.25); }
.kpi b { font-size: 1.05rem; }
hr { margin: 0.9rem 0; }
</style>
""",
    unsafe_allow_html=True,
)

# Build / version banner
with st.sidebar:
    st.caption(f"Build: {APP_BUILD} | tg={getattr(tg, '__version__', '?')}")
    try:
        st.caption(f"tg file: {Path(tg.__file__).name}")
    except Exception:
        pass

DEFAULT_RULES_PATH = Path(__file__).resolve().parent / "Regole_Turni.yml"
DEFAULT_STYLE_TEMPLATE = Path(__file__).resolve().parent / "Style_Template.xlsx"
DEFAULT_UNAV_TEMPLATE = Path(__file__).resolve().parent / "unavailability_template.xlsx"

# ---------------- Secrets helpers ----------------
def _get_secret(path, default=None):
    """Safely read Streamlit secrets with nested keys.

    path: tuple[str, ...] e.g. ("auth","admin_pin") or ("ADMIN_PIN",)
    """
    cur = st.secrets
    for p in path:
        try:
            if isinstance(cur, Mapping) and p in cur:
                cur = cur[p]
            else:
                return default
        except Exception:
            return default
    return cur

def _get_admin_pin() -> str:
    # primary: [auth] admin_pin ; fallback: ADMIN_PIN
    return str(_get_secret(("auth", "admin_pin"), _get_secret(("ADMIN_PIN",), "")) or "")

def _get_doctor_pins() -> dict[str, str]:
    pins = _get_secret(("doctor_pins",), None)
    if isinstance(pins, Mapping):
        return {str(k): str(v) for k, v in pins.items()}
    pins_json = _get_secret(("DOCTOR_PINS_JSON",), "")
    if pins_json:
        try:
            d = yaml.safe_load(pins_json)
            if isinstance(d, Mapping):
                return {str(k): str(v) for k, v in d.items()}
        except Exception:
            pass
    return {}


# ---- Doctor PIN self-service configuration ----
# Email/SMS are optional. Configure at least ONE channel to allow self-service PIN setup/reset.
#
# Secrets format (recommended):
# [smtp]
# host = "smtp.example.com"
# port = 587
# username = "..."
# password = "..."
# from = "turni-utic@example.com"
# starttls = true
#
# [twilio]
# account_sid = "..."
# auth_token = "..."
# from = "+1234567890"
#
# Doctor contacts (who receives OTP) are loaded from GitHub (default: data/doctor_contacts.yml)
# and can contain:
#   Rossi Mario:
#     email: mario.rossi@ospedale.it
#     phone: "+39...."
#
# PIN hashes are stored on GitHub under doctor_auth_dir (default: data/doctor_auth/)

def _smtp_cfg() -> dict:
    cfg = _get_secret(("smtp",), None)
    out: dict = {}
    if isinstance(cfg, Mapping):
        out.update({str(k): cfg[k] for k in cfg.keys()})
    # flat fallbacks
    out.setdefault("host", _get_secret(("SMTP_HOST",), "") or "")
    out.setdefault("port", int(_get_secret(("SMTP_PORT",), 587) or 587))
    out.setdefault("username", _get_secret(("SMTP_USERNAME",), "") or "")
    out.setdefault("password", _get_secret(("SMTP_PASSWORD",), "") or "")
    out.setdefault("from", _get_secret(("SMTP_FROM",), "") or "")
    out.setdefault("starttls", bool(_get_secret(("SMTP_STARTTLS",), True)))
    return out

def _twilio_cfg() -> dict:
    cfg = _get_secret(("twilio",), None)
    out: dict = {}
    if isinstance(cfg, Mapping):
        out.update({str(k): cfg[k] for k in cfg.keys()})
    out.setdefault("account_sid", _get_secret(("TWILIO_ACCOUNT_SID",), "") or "")
    out.setdefault("auth_token", _get_secret(("TWILIO_AUTH_TOKEN",), "") or "")
    out.setdefault("from", _get_secret(("TWILIO_FROM",), "") or "")
    return out

def _email_is_configured() -> bool:
    c = _smtp_cfg()
    return bool(c.get("host") and c.get("from"))

def _sms_is_configured() -> bool:
    c = _twilio_cfg()
    return bool(c.get("account_sid") and c.get("auth_token") and c.get("from"))

def _mask_email(addr: str) -> str:
    s = (addr or "").strip()
    if "@" not in s:
        return s[:2] + "***"
    name, dom = s.split("@", 1)
    if len(name) <= 2:
        name_m = name[:1] + "***"
    else:
        name_m = name[:2] + "***"
    return f"{name_m}@{dom}"

def _mask_phone(p: str) -> str:
    s = re.sub(r"\s+", "", str(p or ""))
    if len(s) <= 4:
        return "***"
    return f"***{s[-3:]}"
def _github_cfg() -> dict:
    cfg = _get_secret(("github_unavailability",), None)
    if isinstance(cfg, Mapping):
        return dict(cfg)
    # fallback flat keys
    return {
        "token": _get_secret(("GITHUB_UNAV_TOKEN",), ""),
        "owner": _get_secret(("GITHUB_UNAV_OWNER",), ""),
        "repo": _get_secret(("GITHUB_UNAV_REPO",), ""),
        "branch": _get_secret(("GITHUB_UNAV_BRANCH",), "main"),
        "path": _get_secret(("GITHUB_UNAV_PATH",), "data/unavailability_store.csv"),
        "settings_path": _get_secret(("GITHUB_UNAV_SETTINGS_PATH",), "data/unavailability_settings.yml"),
        "audit_dir": _get_secret(("GITHUB_UNAV_AUDIT_DIR",), "data/unavailability_audit"),
        "sessions_dir": _get_secret(("GITHUB_UNAV_SESSIONS_DIR",), "data/unavailability_sessions"),
        "doctor_auth_dir": _get_secret(("GITHUB_DOCTOR_AUTH_DIR",), "data/doctor_auth"),
        "contacts_path": _get_secret(("GITHUB_DOCTOR_CONTACTS_PATH",), "data/doctor_contacts.yml"),
        "availability_path": _get_secret(("GITHUB_AVAIL_PATH",), "data/availability_store.csv"),
        "pool_config_path": _get_secret(("GITHUB_POOL_CONFIG_PATH",), "data/pool_config.json"),
    }

# ---------------- Shift history helpers ----------------
def _load_shift_history() -> tuple[dict, str | None]:
    """Carica lo storico turni da GitHub."""
    try:
        sec = st.secrets["github_unavailability"]
        return sh.load_history_from_github(
            sec["owner"], sec["repo"], sec["token"], sec.get("branch", "main"),
        )
    except Exception:
        return {}, None


def _save_shift_history(history: dict, sha: str | None = None) -> bool:
    """Salva lo storico turni su GitHub. Ritorna True se successo."""
    try:
        sec = st.secrets["github_unavailability"]
        # Ricarica lo SHA corrente per evitare conflitti da SHA stale
        try:
            _, current_sha = sh.load_history_from_github(
                sec["owner"], sec["repo"], sec["token"], sec.get("branch", "main"),
            )
        except Exception:
            current_sha = sha
        sh.save_history_to_github(
            history, sec["owner"], sec["repo"], sec["token"],
            sec.get("branch", "main"), current_sha,
        )
        return True
    except Exception as e:
        st.error(f"Errore salvataggio storico: {e}")
        return False


# ---------------- Rules / doctors ----------------
def load_rules_from_source(uploaded) -> tuple[dict, Path]:
    """Return (cfg, rules_path)."""
    if uploaded is None:
        return tg.load_rules(DEFAULT_RULES_PATH), DEFAULT_RULES_PATH
    tmp = Path(tempfile.gettempdir()) / f"rules_{int(time.time())}.yml"
    tmp.write_bytes(uploaded.getvalue())
    return tg.load_rules(tmp), tmp

def doctors_from_cfg(cfg: dict) -> list[str]:
    try:
        return tg.collect_doctors(cfg)
    except Exception:
        return sorted(set((cfg.get("doctors") or [])))

# ---------------- GitHub datastore ops ----------------

def _unavail_per_doctor_dir() -> str:
    """GitHub directory path for per-doctor unavailability CSV files."""
    g = _github_cfg()
    return str(g.get("per_doctor_dir") or "data/unavailability").rstrip("/")


def _doctor_unavail_path(doctor: str) -> str:
    """Full GitHub path for a single doctor's unavailability CSV."""
    return f"{_unavail_per_doctor_dir()}/unavail_{_doctor_slug(doctor)}.csv"


def load_doctor_unavail_from_github(doctor: str) -> tuple[list[dict], str | None]:
    """Load ONLY the given doctor's unavailability rows from their personal CSV file."""
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo")):
        raise RuntimeError("GitHub non configurato (unavailability store).")
    gf = github_utils.get_file(
        owner=g["owner"],
        repo=g["repo"],
        path=_doctor_unavail_path(doctor),
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    if gf is None:
        return [], None
    return ustore.load_store(gf.text), gf.sha


def save_doctor_unavail_to_github(
    doctor: str,
    rows: list[dict],
    sha: str | None,
    message: str,
) -> str | None:
    """Write ONLY the given doctor's rows to their personal CSV file on GitHub."""
    g = _github_cfg()
    text = ustore.to_csv(rows)
    resp = github_utils.put_file(
        owner=g["owner"],
        repo=g["repo"],
        path=_doctor_unavail_path(doctor),
        token=g["token"],
        branch=g.get("branch", "main"),
        sha=sha,
        message=message,
        text=text,
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content["sha"])
    except Exception:
        pass
    return None


def load_store_from_github() -> tuple[list[dict], str | None]:
    """Load all unavailability rows.

    Primary source: per-doctor CSV files in the per_doctor_dir directory.
    Each doctor saves to their own file — no cross-doctor race conditions.
    Fallback: legacy aggregate CSV (path key in secrets) if per-doctor
    directory is empty or does not yet exist (pre-migration).

    Returns (rows, sha_or_None). SHA is None when aggregating multiple files.
    """
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo")):
        raise RuntimeError("Archivio indisponibilità: secrets GitHub non configurati.")

    per_dir = _unavail_per_doctor_dir()
    files = github_utils.list_dir(
        owner=g["owner"],
        repo=g["repo"],
        path=per_dir,
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    csv_files = [f for f in files if f.get("name", "").endswith(".csv")]

    if csv_files:
        all_rows: list[dict] = []
        for file_meta in csv_files:
            gf = github_utils.get_file(
                owner=g["owner"],
                repo=g["repo"],
                path=file_meta["path"],
                token=g["token"],
                branch=g.get("branch", "main"),
            )
            if gf:
                all_rows.extend(ustore.load_store(gf.text))
        return all_rows, None  # no single SHA represents the full aggregate

    # Fallback: legacy single-file CSV
    if not g.get("path"):
        return [], None
    gf = github_utils.get_file(
        owner=g["owner"],
        repo=g["repo"],
        path=g["path"],
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    if gf is None:
        return [], None
    return ustore.load_store(gf.text), gf.sha


def save_store_to_github(rows: list[dict], sha: str | None, message: str) -> str | None:
    g = _github_cfg()
    if not g.get("path"):
        raise RuntimeError("Chiave 'path' non configurata in github_unavailability secrets (legacy store).")
    text = ustore.to_csv(rows)
    resp = github_utils.put_file(
        owner=g["owner"],
        repo=g["repo"],
        path=g["path"],
        token=g["token"],
        branch=g.get("branch", "main"),
        sha=sha,
        message=message,
        text=text,
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content.get("sha"))
    except Exception:
        pass
    return None


# ── Per-doctor availability file helpers ───────────────────────────────────

def _avail_per_doctor_dir() -> str:
    """GitHub directory path for per-doctor availability CSV files."""
    g = _github_cfg()
    return str(g.get("per_doctor_avail_dir") or "data/availability").rstrip("/")


def _doctor_avail_path(doctor: str) -> str:
    """Full GitHub path for a single doctor's availability CSV."""
    return f"{_avail_per_doctor_dir()}/avail_{_doctor_slug(doctor)}.csv"


def load_doctor_avail_from_github(doctor: str) -> tuple[list[dict], str | None]:
    """Load ONLY the given doctor's availability rows from their personal CSV file."""
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo")):
        raise RuntimeError("GitHub non configurato (availability store).")
    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"],
        path=_doctor_avail_path(doctor),
        token=g["token"], branch=g.get("branch", "main"),
    )
    if gf is None:
        return [], None
    return ustore.load_store(gf.text), gf.sha


def save_doctor_avail_to_github(
    doctor: str,
    rows: list[dict],
    sha: str | None,
    message: str,
) -> str | None:
    """Write ONLY the given doctor's rows to their personal availability CSV file."""
    g = _github_cfg()
    text = ustore.to_csv(rows)
    resp = github_utils.put_file(
        owner=g["owner"], repo=g["repo"],
        path=_doctor_avail_path(doctor),
        token=g["token"], branch=g.get("branch", "main"),
        sha=sha, message=message, text=text,
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content["sha"])
    except Exception:
        pass
    return None


def load_avail_store_from_github() -> tuple[list[dict], str | None]:
    """Load all availability rows.

    Primary source: per-doctor CSV files in per_doctor_avail_dir.
    Fallback: legacy aggregate CSV (availability_path key in secrets).
    Returns (rows, sha_or_None). SHA is None when aggregating multiple files.
    """
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo")):
        raise RuntimeError("GitHub non configurato (availability store).")

    per_dir = _avail_per_doctor_dir()
    files = github_utils.list_dir(
        owner=g["owner"], repo=g["repo"], path=per_dir,
        token=g["token"], branch=g.get("branch", "main"),
    )
    csv_files = [f for f in files if f.get("name", "").endswith(".csv")]

    if csv_files:
        all_rows: list[dict] = []
        for file_meta in csv_files:
            gf = github_utils.get_file(
                owner=g["owner"], repo=g["repo"], path=file_meta["path"],
                token=g["token"], branch=g.get("branch", "main"),
            )
            if gf:
                all_rows.extend(ustore.load_store(gf.text))
        return all_rows, None

    # Fallback: legacy single-file CSV
    path = g.get("availability_path", "data/availability_store.csv")
    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"], path=path,
        token=g["token"], branch=g.get("branch", "main"),
    )
    if gf is None:
        return [], None
    return ustore.load_store(gf.text), gf.sha


def save_avail_store_to_github(rows: list[dict], sha: str | None, message: str) -> str | None:
    """Legacy writer — kept for backward compatibility. Prefer save_doctor_avail_to_github."""
    g = _github_cfg()
    path = g.get("availability_path", "data/availability_store.csv")
    text = ustore.to_csv(rows)
    resp = github_utils.put_file(
        owner=g["owner"], repo=g["repo"], path=path,
        token=g["token"], branch=g.get("branch", "main"),
        sha=sha, message=message, text=text,
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content.get("sha"))
    except Exception:
        pass
    return None


def save_doctor_availability_with_retry(
    *,
    doctor: str,
    entries_by_month: dict,
    updated_at: str,
    message: str,
    initial_rows: list[dict] | None = None,
    initial_sha: str | None = None,
    max_retries: int = 6,
) -> str | None:
    """Concurrency-safe save using per-doctor availability CSV files.

    Each doctor writes only their own file — no cross-doctor SHA conflicts.
    Returns: final_sha (or None)
    """
    last_err: Exception | None = None
    months = sorted(entries_by_month.items())

    for attempt in range(max_retries):
        if attempt == 0 and initial_rows is not None:
            doctor_rows = [r for r in initial_rows if r.get("doctor", "") == doctor]
            doctor_sha = initial_sha
        else:
            doctor_rows, doctor_sha = load_doctor_avail_from_github(doctor)

        new_rows = list(doctor_rows)
        for (yy, mm), entries in months:
            new_rows = ustore.replace_doctor_month(
                new_rows, doctor, int(yy), int(mm), entries, updated_at=updated_at
            )

        try:
            _new_sha = save_doctor_avail_to_github(doctor, new_rows, doctor_sha, message)
            verified_rows, latest_sha = load_doctor_avail_from_github(doctor)
            return latest_sha or _new_sha
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                sleep_s = min(3.0, 0.35 * (2 ** attempt) + random.random() * 0.25)
                time.sleep(sleep_s)
                continue
            raise

    if last_err:
        raise last_err
    raise RuntimeError("Errore salvataggio preferenze: tentativi esauriti.")


def _is_sha_conflict_error(err: Exception) -> bool:
    """Return True if the HTTP error likely indicates a concurrent update (SHA mismatch / conflict)."""
    if isinstance(err, requests.HTTPError):
        resp = getattr(err, "response", None)
        if resp is None:
            return False
        code = getattr(resp, "status_code", None)
        if code in (409, 412):
            return True
        if code == 422:
            # GitHub Contents API sometimes returns 422 for a SHA mismatch
            try:
                j = resp.json() if hasattr(resp, "json") else {}
                msg = str(j.get("message", "") or "").lower()
            except Exception:
                msg = str(getattr(resp, "text", "") or "").lower()
            if "sha" in msg and ("match" in msg or "invalid" in msg or "does not" in msg):
                return True
        return False
    return False


# ---------------- Doctor session lease (GitHub) ----------------
def _doctor_slug(doctor: str) -> str:
    """Filesystem-like slug for a doctor name."""
    s = (doctor or "").strip().lower()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_\-]", "", s)
    return s or "doctor"



# ---------------- Doctor PIN store (GitHub) ----------------
# Storage model (per doctor file, to avoid cross-doctor conflicts):
#   <doctor_auth_dir>/pin_<doctor>.json
#   <doctor_auth_dir>/otp_<doctor>.json   (temporary OTP for setup/reset)
#
# Each PIN is stored ONLY as a salted PBKDF2 hash (never plaintext).

PIN_PBKDF2_ITERS = 200_000
OTP_TTL_MINUTES = 10
OTP_MAX_ATTEMPTS = 5

def _doctor_auth_dir() -> str:
    g = _github_cfg()
    return str((g.get("doctor_auth_dir") or "data/doctor_auth")).rstrip("/")

def _doctor_pin_path(doctor: str) -> str:
    return f"{_doctor_auth_dir()}/pin_{_doctor_slug(doctor)}.json"

def _doctor_otp_path(doctor: str) -> str:
    return f"{_doctor_auth_dir()}/otp_{_doctor_slug(doctor)}.json"

def _b64e(b: bytes) -> str:
    return base64.b64encode(b).decode("ascii")

def _b64d(s: str) -> bytes:
    return base64.b64decode(s.encode("ascii"))

def _hash_pin(pin: str, salt: bytes | None = None, iters: int = PIN_PBKDF2_ITERS) -> tuple[str, str, int]:
    pin_b = str(pin or "").encode("utf-8")
    if salt is None:
        salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", pin_b, salt, int(iters))
    return _b64e(salt), _b64e(dk), int(iters)

def _verify_pin(pin: str, rec: dict) -> bool:
    try:
        salt = _b64d(str(rec.get("salt_b64") or ""))
        want = _b64d(str(rec.get("hash_b64") or ""))
        iters = int(rec.get("iters") or PIN_PBKDF2_ITERS)
    except Exception:
        return False
    dk = hashlib.pbkdf2_hmac("sha256", str(pin or "").encode("utf-8"), salt, iters)
    return hmac.compare_digest(dk, want)

@st.cache_data(ttl=120)
def load_doctor_contacts_from_github() -> dict:
    """Return {doctor: {email, phone}} from a YAML file.

    Primary source: GitHub repo configured in secrets (same repo used for the unavailability store).
    Fallback: local file inside the deployed app repo (useful if you keep contacts in the code repo).
    """
    g = _github_cfg()
    path = g.get("contacts_path") or "data/doctor_contacts.yml"

    text: str | None = None
    gf = None
    try:
        gf = github_utils.get_file(
            owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch", "main")
        )
    except Exception:
        gf = None

    if gf is not None and isinstance(getattr(gf, "text", None), str):
        text = gf.text
    else:
        # local fallback (Streamlit Cloud has the repo checked out on disk)
        try:
            lp = Path(path)
            if lp.exists() and lp.is_file():
                text = lp.read_text(encoding="utf-8", errors="replace")
        except Exception:
            text = None

    if not text:
        return {}

    try:
        data = yaml.safe_load(text) or {}
    except Exception:
        data = {}
    if not isinstance(data, Mapping):
        return {}
    out: dict = {}
    for k, v in data.items():
        if not isinstance(v, Mapping):
            continue
        out[str(k)] = {"email": str(v.get("email") or ""), "phone": str(v.get("phone") or "")}
    return out


def _doctor_key_norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (str(s or "")).casefold())

def get_doctor_contact(doctor: str) -> dict:
    """Robust contact lookup (exact / casefold / normalized)."""
    contacts = load_doctor_contacts_from_github()
    dk = str(doctor or "").strip()
    if dk in contacts:
        return contacts[dk] or {}
    dkl = dk.casefold()
    for k, v in contacts.items():
        if str(k).strip().casefold() == dkl:
            return v or {}
    dkn = _doctor_key_norm(dk)
    if dkn:
        for k, v in contacts.items():
            if _doctor_key_norm(k) == dkn:
                return v or {}
    return {}

def load_doctor_pin_record(doctor: str) -> tuple[dict | None, str | None]:
    """Load the per-doctor PIN record file from GitHub."""
    g = _github_cfg()
    path = _doctor_pin_path(doctor)
    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch","main")
    )
    if gf is None:
        return None, None
    try:
        rec = json.loads(gf.text or "{}")
        if isinstance(rec, dict):
            return rec, gf.sha
    except Exception:
        pass
    return None, gf.sha

def save_doctor_pin_record(doctor: str, rec: dict, sha: str | None, message: str) -> str | None:
    g = _github_cfg()
    path = _doctor_pin_path(doctor)
    text = json.dumps(rec, ensure_ascii=False, indent=2)
    resp = github_utils.put_file(
        owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch","main"),
        sha=sha, message=message, text=text
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content.get("sha"))
    except Exception:
        pass
    return None

def set_doctor_pin_with_retry(doctor: str, new_pin: str, reason: str) -> None:
    """Set/update the doctor PIN (hash) with optimistic concurrency."""
    last_err: Exception | None = None
    for attempt in range(6):
        rec, sha = load_doctor_pin_record(doctor)
        rec = dict(rec or {})
        salt_b64, hash_b64, iters = _hash_pin(new_pin)
        rec.update({
            "doctor": (doctor or "").strip(),
            "salt_b64": salt_b64,
            "hash_b64": hash_b64,
            "iters": iters,
            "pin_updated_at": _utc_now_iso(),
            "app_build": APP_BUILD,
        })
        try:
            save_doctor_pin_record(doctor, rec, sha, message=f"Set PIN {doctor}: {reason}")
            return
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                time.sleep(min(1.6, 0.25 * (2 ** attempt) + random.random() * 0.2))
                continue
            raise
    if last_err:
        raise last_err
    raise RuntimeError("Errore impostazione PIN: tentativi esauriti.")

def verify_doctor_pin(doctor: str, pin: str) -> bool:
    """Verify PIN against GitHub record; fallback to secrets doctor_pins for migration."""
    rec, _sha = load_doctor_pin_record(doctor)
    if isinstance(rec, dict) and rec.get("hash_b64") and rec.get("salt_b64"):
        return _verify_pin(pin, rec)
    # migration fallback (old secrets-based PINs)
    pins = _get_doctor_pins()
    expected = str(pins.get(doctor, ""))
    return bool(pin) and bool(expected) and (pin == expected)

def doctor_has_pin(doctor: str) -> bool:
    rec, _sha = load_doctor_pin_record(doctor)
    return bool(isinstance(rec, dict) and rec.get("hash_b64") and rec.get("salt_b64"))


def _send_otp_email(dest: str, code: str) -> None:
    cfg = _smtp_cfg()
    if not _email_is_configured():
        raise RuntimeError("Invio email non configurato (smtp).")
    msg = EmailMessage()
    msg["Subject"] = "Codice verifica – Turni UTIC"
    msg["From"] = cfg.get("from")
    msg["To"] = dest
    msg.set_content(
        "Hai richiesto un codice per impostare o resettare il PIN di accesso.\n\n"
        f"CODICE: {code}\n\n"
        "Se non sei stato tu, ignora questo messaggio."
    )
    host = str(cfg.get("host") or "")
    port = int(cfg.get("port") or 587)
    username = str(cfg.get("username") or "")
    password = str(cfg.get("password") or "")
    starttls = bool(cfg.get("starttls", True))
    with smtplib.SMTP(host, port, timeout=20) as s:
        s.ehlo()
        if starttls:
            s.starttls()
            s.ehlo()
        if username and password:
            s.login(username, password)
        s.send_message(msg)

def _send_otp_sms(dest: str, code: str) -> None:
    cfg = _twilio_cfg()
    if not _sms_is_configured():
        raise RuntimeError("Invio SMS non configurato (twilio).")
    sid = str(cfg.get("account_sid") or "")
    token = str(cfg.get("auth_token") or "")
    from_num = str(cfg.get("from") or "")
    url = f"https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json"
    data = {
        "From": from_num,
        "To": dest,
        "Body": f"Turni UTIC - codice verifica PIN: {code}",
    }
    r = requests.post(url, data=data, auth=(sid, token), timeout=20)
    if r.status_code >= 400:
        raise RuntimeError(f"Errore invio SMS (Twilio): {r.status_code} {r.text[:200]}")

def load_doctor_otp_record(doctor: str) -> tuple[dict | None, str | None]:
    g = _github_cfg()
    path = _doctor_otp_path(doctor)
    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch","main")
    )
    if gf is None:
        return None, None
    try:
        rec = json.loads(gf.text or "{}")
        if isinstance(rec, dict):
            return rec, gf.sha
    except Exception:
        pass
    return None, gf.sha

def save_doctor_otp_record(doctor: str, rec: dict, sha: str | None, message: str) -> str | None:
    g = _github_cfg()
    path = _doctor_otp_path(doctor)
    text = json.dumps(rec, ensure_ascii=False, indent=2)
    resp = github_utils.put_file(
        owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch","main"),
        sha=sha, message=message, text=text
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content.get("sha"))
    except Exception:
        pass
    return None

def request_pin_otp(doctor: str, channel: str) -> str:
    """Send an OTP to the doctor's configured email/SMS. Returns masked destination."""
    contacts = load_doctor_contacts_from_github()
    c = contacts.get(doctor) or {}
    email = str(c.get("email") or "").strip()
    phone = str(c.get("phone") or "").strip()

    channel = str(channel or "").strip().lower()
    if channel == "email":
        if not email:
            raise RuntimeError("Email non configurata per questo medico.")
        if not _email_is_configured():
            raise RuntimeError("Invio email non disponibile: configura SMTP in secrets.")
        dest = email
        masked = _mask_email(dest)
        sender = _send_otp_email
    elif channel == "sms":
        if not phone:
            raise RuntimeError("Numero di telefono non configurato per questo medico.")
        if not _sms_is_configured():
            raise RuntimeError("Invio SMS non disponibile: configura Twilio in secrets.")
        dest = phone
        masked = _mask_phone(dest)
        sender = _send_otp_sms
    else:
        raise RuntimeError("Canale OTP non valido.")

    # generate code
    code = f"{random.randint(0, 999999):06d}"
    salt = os.urandom(16)
    code_hash = hashlib.pbkdf2_hmac("sha256", code.encode("utf-8"), salt, 120_000)

    expires_dt = datetime.utcnow() + timedelta(minutes=OTP_TTL_MINUTES)

    last_err: Exception | None = None
    for attempt in range(6):
        rec, sha = load_doctor_otp_record(doctor)
        rec = dict(rec or {})
        rec.update({
            "doctor": (doctor or "").strip(),
            "channel": channel,
            "dest_masked": masked,
            "salt_b64": _b64e(salt),
            "code_hash_b64": _b64e(code_hash),
            "created_at": _utc_now_iso(),
            "expires_at": expires_dt.isoformat(timespec="seconds") + "Z",
            "attempts": 0,
            "max_attempts": OTP_MAX_ATTEMPTS,
            "app_build": APP_BUILD,
            "used_at": None,
        })
        try:
            save_doctor_otp_record(doctor, rec, sha, message=f"OTP request {doctor} ({channel})")
            break
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                time.sleep(min(1.4, 0.25 * (2 ** attempt) + random.random() * 0.2))
                continue
            raise
    else:
        if last_err:
            raise last_err
        raise RuntimeError("Errore OTP: tentativi esauriti.")

    # send AFTER persisting record (so we don't send a code we can't verify)
    sender(dest, code)
    return masked

def verify_pin_otp_and_consume(doctor: str, code: str) -> None:
    """Verify OTP; on success marks it as used (prevents replay)."""
    code = str(code or "").strip()
    if not re.fullmatch(r"\d{6}", code):
        raise RuntimeError("Codice non valido (deve essere di 6 cifre).")

    last_err: Exception | None = None
    for attempt in range(6):
        rec, sha = load_doctor_otp_record(doctor)
        if not isinstance(rec, dict):
            raise RuntimeError("Nessun codice attivo. Richiedi un nuovo codice.")
        exp = _parse_utc_iso(rec.get("expires_at"))
        now_utc = datetime.utcnow()
        if exp is None or (exp.tzinfo and exp.astimezone(timezone.utc).replace(tzinfo=None) <= now_utc) or ((not exp.tzinfo) and exp <= now_utc):
            raise RuntimeError("Codice scaduto. Richiedi un nuovo codice.")
        if rec.get("used_at"):
            raise RuntimeError("Codice già utilizzato. Richiedi un nuovo codice.")
        attempts = int(rec.get("attempts") or 0)
        max_a = int(rec.get("max_attempts") or OTP_MAX_ATTEMPTS)
        if attempts >= max_a:
            raise RuntimeError("Troppi tentativi. Richiedi un nuovo codice.")

        try:
            salt = _b64d(str(rec.get("salt_b64") or ""))
            want = _b64d(str(rec.get("code_hash_b64") or ""))
        except Exception:
            raise RuntimeError("Codice non disponibile. Richiedi un nuovo codice.")

        got = hashlib.pbkdf2_hmac("sha256", code.encode("utf-8"), salt, 120_000)
        ok = hmac.compare_digest(got, want)

        rec2 = dict(rec)
        rec2["attempts"] = attempts + (0 if ok else 1)
        if ok:
            rec2["used_at"] = _utc_now_iso()

        try:
            save_doctor_otp_record(doctor, rec2, sha, message=f"OTP verify {doctor}")
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                time.sleep(min(1.4, 0.25 * (2 ** attempt) + random.random() * 0.2))
                continue
            raise

        if ok:
            return
        raise RuntimeError("Codice errato. Riprova.")

    if last_err:
        raise last_err
    raise RuntimeError("Errore OTP: tentativi esauriti.")


def _session_lease_path(doctor: str) -> str:
    g = _github_cfg()
    sessions_dir = (g.get("sessions_dir") or "data/unavailability_sessions").rstrip("/")
    return f"{sessions_dir}/lease_{_doctor_slug(doctor)}.json"


def _parse_utc_iso(ts: str) -> datetime | None:
    s = str(ts or "").strip()
    if not s:
        return None
    # support ...Z
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def load_session_lease_from_github(doctor: str) -> tuple[dict | None, str | None]:
    g = _github_cfg()
    path = _session_lease_path(doctor)
    gf = github_utils.get_file(
        owner=g["owner"],
        repo=g["repo"],
        path=path,
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    if gf is None:
        return None, None
    try:
        data = json.loads(gf.text or "{}")
        if isinstance(data, dict):
            return data, gf.sha
    except Exception:
        pass
    return None, gf.sha


def save_session_lease_to_github(doctor: str, lease: dict, sha: str | None, message: str) -> str | None:
    g = _github_cfg()
    path = _session_lease_path(doctor)
    text = json.dumps(lease, ensure_ascii=False, indent=2)
    resp = github_utils.put_file(
        owner=g["owner"],
        repo=g["repo"],
        path=path,
        token=g["token"],
        branch=g.get("branch", "main"),
        sha=sha,
        message=message,
        text=text,
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content.get("sha"))
    except Exception:
        pass
    return None


def _lease_is_expired(lease: dict | None, now_utc: datetime) -> bool:
    if not isinstance(lease, dict):
        return True
    exp = _parse_utc_iso(lease.get("expires_at"))
    if exp is not None:
        try:
            exp_naive = exp.astimezone(timezone.utc).replace(tzinfo=None) if exp.tzinfo else exp
        except Exception:
            exp_naive = exp.replace(tzinfo=None)
        return exp_naive <= now_utc
    # fallback: last_seen + TTL
    last_seen = _parse_utc_iso(lease.get("last_seen") or lease.get("issued_at"))
    if last_seen is None:
        return True
    try:
        age_s = (now_utc - (last_seen.astimezone(timezone.utc).replace(tzinfo=None) if last_seen.tzinfo else last_seen)).total_seconds()
    except Exception:
        return True
    return age_s > (DOCTOR_SESSION_TTL_MINUTES * 60)


def acquire_doctor_session_lease(
    *,
    doctor: str,
    session_id: str,
    max_retries: int = 6,
) -> tuple[dict, str | None]:
    """Acquire/overwrite the lease for this doctor (kicking out other sessions).

    Uses optimistic concurrency with retries on SHA conflicts.
    """
    last_err: Exception | None = None
    for attempt in range(max_retries):
        lease, sha = load_session_lease_from_github(doctor)

        now_utc = datetime.utcnow()
        expires_at = (now_utc.timestamp() + DOCTOR_SESSION_TTL_MINUTES * 60)
        expires_dt = datetime.utcfromtimestamp(expires_at)

        new_lease = {
            "doctor": (doctor or "").strip(),
            "session_id": session_id,
            "issued_at": lease.get("issued_at") if isinstance(lease, dict) and lease.get("session_id") == session_id else _utc_now_iso(),
            "last_seen": _utc_now_iso(),
            "expires_at": expires_dt.isoformat(timespec="seconds") + "Z",
            "app_build": APP_BUILD,
        }

        try:
            new_sha = save_session_lease_to_github(
                doctor,
                new_lease,
                sha,
                message=f"Lease doctor session: {doctor} ({new_lease['last_seen']})",
            )
            return new_lease, new_sha
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                time.sleep(min(1.6, 0.25 * (2 ** attempt) + random.random() * 0.2))
                continue
            raise

    if last_err:
        raise last_err
    raise RuntimeError("Errore lease sessione medico: tentativi esauriti.")


def check_doctor_session_lease(doctor: str, session_id: str) -> bool:
    """Return True if the current lease belongs to this session (and is not expired)."""
    lease, _sha = load_session_lease_from_github(doctor)
    now_utc = datetime.utcnow()
    if _lease_is_expired(lease, now_utc):
        # Treat expired as "free": current session can re-acquire on next action.
        return True
    if not isinstance(lease, dict):
        return True
    return str(lease.get("session_id") or "") == str(session_id or "")


def _month_entries_signature(rows: list[dict]) -> list[tuple[str, str, str]]:
    """Return a deterministic signature for month rows: (date, shift, note)."""
    sig: set[tuple[str, str, str]] = set()
    for r in rows:
        try:
            d_iso = ustore.parse_iso_date(r.get("date", "")).isoformat()
        except Exception:
            continue
        sh = ustore.norm_shift(r.get("shift", ""))
        if not sh:
            continue
        sig.add((d_iso, sh, str(r.get("note", "") or "")))
    return sorted(sig)


def _entries_signature_from_tuples(entries: list[tuple[date, str, str]]) -> list[tuple[str, str, str]]:
    """Signature for a list of (date, shift, note)."""
    sig: set[tuple[str, str, str]] = set()
    for d, sh, note in (entries or []):
        if not isinstance(d, date):
            continue
        sh2 = ustore.norm_shift(sh)
        if not sh2:
            continue
        sig.add((d.isoformat(), sh2, str(note or "")))
    return sorted(sig)


def _build_expected_signatures(
    rows: list[dict],
    doctor: str,
    months: list[tuple[int, int]],
) -> dict[tuple[int, int], list[tuple[str, str, str]]]:
    """Build signatures as seen by the editor when doctor data is initially loaded."""
    out: dict[tuple[int, int], list[tuple[str, str, str]]] = {}
    for yy, mm in months:
        existing = ustore.filter_doctor_month(rows, doctor, int(yy), int(mm))
        out[(int(yy), int(mm))] = _month_entries_signature(existing)
    return out


def _detect_stale_doctor_month(
    rows: list[dict],
    doctor: str,
    expected_signatures: dict[tuple[int, int], list[tuple[str, str, str]]],
) -> str | None:
    """Return stale month label (YYYY-MM) if persisted data changed after load."""
    for (yy, mm), expected in sorted(expected_signatures.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        current_rows = ustore.filter_doctor_month(rows, doctor, int(yy), int(mm))
        if _month_entries_signature(current_rows) != expected:
            return f"{yy}-{mm:02d}"
    return None


def save_doctor_unavailability_with_retry(
    *,
    doctor: str,
    normalized_entries_by_month: dict[tuple[int, int], list[tuple[date, str, str]]],
    updated_at: str,
    message: str,
    initial_rows: list[dict] | None = None,
    initial_sha: str | None = None,
    expected_signatures: dict[tuple[int, int], list[tuple[str, str, str]]] | None = None,
    max_retries: int = 6,
) -> tuple[list[tuple[str, dict]], str | None]:
    """Concurrency-safe save using per-doctor CSV files.

    Each doctor writes only their own file — no cross-doctor SHA conflicts.
    The only conflict scenario is the same doctor saving from two browser
    tabs simultaneously, which the session lease already prevents.

    Returns: (audit_todo, final_sha)
    """
    last_err: Exception | None = None
    months = sorted(normalized_entries_by_month.items(), key=lambda kv: (kv[0][0], kv[0][1]))

    for attempt in range(max_retries):
        # First attempt: reuse the rows already loaded at page render time
        # (filtered to this doctor) to avoid an extra round-trip.
        if attempt == 0 and initial_rows is not None:
            doctor_rows = [r for r in initial_rows if r.get("doctor", "") == doctor]
            doctor_sha = initial_sha
        else:
            doctor_rows, doctor_sha = load_doctor_unavail_from_github(doctor)

        new_rows = list(doctor_rows)
        audit_todo: list[tuple[str, dict]] = []

        for (yy, mm), entries_norm in months:
            yy_i, mm_i = int(yy), int(mm)
            existing_rows = ustore.filter_doctor_month(doctor_rows, doctor, yy_i, mm_i)
            diff = compute_unavailability_diff(existing_rows, entries_norm)
            if diff.get("added_count") or diff.get("removed_count") or diff.get("note_changed_count"):
                audit_todo.append((f"{yy_i}-{mm_i:02d}", diff))
            new_rows = ustore.replace_doctor_month(
                new_rows, doctor, yy_i, mm_i, entries_norm, updated_at=updated_at
            )

        try:
            _new_sha = save_doctor_unavail_to_github(doctor, new_rows, doctor_sha, message)

            # Read-back verification: confirm what's on GitHub matches what we wrote.
            verified_rows, latest_sha = load_doctor_unavail_from_github(doctor)
            for (yy, mm), entries_norm in months:
                yy_i, mm_i = int(yy), int(mm)
                persisted = ustore.filter_doctor_month(verified_rows, doctor, yy_i, mm_i)
                if _month_entries_signature(persisted) != _entries_signature_from_tuples(entries_norm):
                    raise RuntimeError(
                        "Salvataggio non verificato: i dati sul server non corrispondono a quanto inserito. "
                        "Ricarica e riprova."
                    )

            return audit_todo, (latest_sha or _new_sha)
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                sleep_s = min(3.0, 0.35 * (2 ** attempt) + random.random() * 0.25)
                time.sleep(sleep_s)
                continue
            # Retry anche per lag GitHub nella verifica post-save
            if "non verificato" in str(e) and attempt < max_retries - 1:
                time.sleep(min(2.0, 0.5 * (attempt + 1)))
                continue
            raise

    if last_err:
        raise last_err
    raise RuntimeError("Errore salvataggio: tentativi esauriti senza dettaglio.")

# ---------------- GitHub settings & audit log ----------------
DEFAULT_SETTINGS = {
    "unavailability_open": True,
    "max_unavailability_per_shift": 6,
    "max_availability_per_shift": 6,  # max preferenze disponibilità per fascia per mese
    "max_weekend_days": MAX_WEEKEND_DAYS,  # max sabati e max domeniche distinti per mese
    "doctor_caps": {},  # cap personalizzato per medico: {"Dattilo": 10, "De Gregorio": 10, "Zito": 10}
}

AUDIT_FIELDS = [
    "ts_utc",
    "doctor",
    "month",
    "action",
    "before_count",
    "after_count",
    "added_count",
    "removed_count",
    "note_changed_count",
    "details_json",
    "app_build",
]

def load_app_settings_from_github() -> tuple[dict, str | None]:
    """Load app settings (toggle unavailability entry + max per shift) from GitHub.

    If the settings file doesn't exist yet, returns defaults.
    """
    g = _github_cfg()
    path = g.get("settings_path") or "data/unavailability_settings.yml"
    gf = github_utils.get_file(
        owner=g["owner"],
        repo=g["repo"],
        path=path,
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    if gf is None:
        # Defaults (no file yet)
        return dict(DEFAULT_SETTINGS), None

    try:
        data = yaml.safe_load(gf.text) or {}
    except Exception:
        data = {}

    if not isinstance(data, Mapping):
        data = {}

    out = dict(DEFAULT_SETTINGS)

    # allow some legacy key names
    if "unavailability_open" in data:
        out["unavailability_open"] = bool(data.get("unavailability_open"))
    elif "unavailability_enabled" in data:
        out["unavailability_open"] = bool(data.get("unavailability_enabled"))
    elif "open" in data:
        out["unavailability_open"] = bool(data.get("open"))

    try:
        out["max_unavailability_per_shift"] = int(
            data.get("max_unavailability_per_shift", data.get("max_per_shift", DEFAULT_SETTINGS["max_unavailability_per_shift"]))
        )
    except Exception:
        out["max_unavailability_per_shift"] = DEFAULT_SETTINGS["max_unavailability_per_shift"]

    try:
        out["max_availability_per_shift"] = int(
            data.get("max_availability_per_shift", DEFAULT_SETTINGS["max_availability_per_shift"])
        )
    except Exception:
        out["max_availability_per_shift"] = DEFAULT_SETTINGS["max_availability_per_shift"]

    try:
        out["max_weekend_days"] = int(
            data.get("max_weekend_days", DEFAULT_SETTINGS["max_weekend_days"])
        )
    except Exception:
        out["max_weekend_days"] = DEFAULT_SETTINGS["max_weekend_days"]

    # cap personalizzati per medico
    try:
        dc = data.get("doctor_caps", {})
        out["doctor_caps"] = {str(k): int(v) for k, v in (dc or {}).items()} if isinstance(dc, dict) else {}
    except Exception:
        out["doctor_caps"] = {}

    # optional metadata
    out["updated_at"] = str(data.get("updated_at") or "")
    out["updated_by"] = str(data.get("updated_by") or "")

    # defensive bounds
    if out["max_unavailability_per_shift"] < 0:
        out["max_unavailability_per_shift"] = 0
    if out["max_weekend_days"] < 0:
        out["max_weekend_days"] = 0

    return out, gf.sha

def save_app_settings_to_github(settings: dict, sha: str | None, message: str):
    g = _github_cfg()
    path = g.get("settings_path") or "data/unavailability_settings.yml"
    # Write as YAML for readability
    text = yaml.safe_dump(settings, sort_keys=False, allow_unicode=True)
    github_utils.put_file(
        owner=g["owner"],
        repo=g["repo"],
        path=path,
        token=g["token"],
        branch=g.get("branch", "main"),
        sha=sha,
        message=message,
        text=text,
    )

def load_pool_config_from_github_st() -> tuple[dict, str | None]:
    """Carica pool_config.json da GitHub. Ritorna ({}, None) se assente."""
    import pool_config_store as _pcs
    g = _github_cfg()
    path = g.get("pool_config_path") or _pcs.POOL_CONFIG_PATH_DEFAULT
    return _pcs.load_pool_config_from_github(
        owner=g["owner"],
        repo=g["repo"],
        token=g["token"],
        branch=g.get("branch", "main"),
        path=path,
    )


def save_pool_config_with_retry(cfg: dict, sha: str | None, max_retries: int = 3) -> tuple[bool, str]:
    """Salva pool_config su GitHub con retry in caso di conflitto SHA.

    Ritorna (ok: bool, message: str).
    """
    import pool_config_store as _pcs
    g = _github_cfg()
    path = g.get("pool_config_path") or _pcs.POOL_CONFIG_PATH_DEFAULT
    current_sha = sha
    for attempt in range(max_retries):
        try:
            _pcs.save_pool_config_to_github(
                cfg=cfg,
                owner=g["owner"],
                repo=g["repo"],
                token=g["token"],
                branch=g.get("branch", "main"),
                sha=current_sha,
                path=path,
            )
            return True, "Configurazione pool salvata."
        except Exception as e:
            if _is_sha_conflict_error(e) and attempt < max_retries - 1:
                # Ricarica SHA aggiornato e riprova
                fresh, fresh_sha = load_pool_config_from_github_st()
                current_sha = fresh_sha
                continue
            return False, f"Errore salvataggio pool config: {e}"
    return False, "Impossibile salvare dopo i tentativi massimi."


def _audit_path_for_month(mk: str) -> str:
    g = _github_cfg()
    audit_dir = g.get("audit_dir") or "data/unavailability_audit"
    return f"{audit_dir}/unavailability_audit_{mk}.csv"


@st.cache_data(ttl=60)
def load_audit_log_text_from_github(mk: str) -> str | None:
    """Return monthly audit log CSV text (or None if missing)."""
    g = _github_cfg()
    path = _audit_path_for_month(mk)
    gf = github_utils.get_file(
        owner=g["owner"],
        repo=g["repo"],
        path=path,
        token=g["token"],
        branch=g.get("branch", "main"),
    )
    return gf.text if gf else None


def audit_df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "audit") -> bytes:
    """Convert an audit dataframe to an .xlsx in-memory."""
    buf = io.BytesIO()
    # Use openpyxl engine (already in requirements)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31] or "audit")
    return buf.getvalue()

def _entries_map_from_store_rows(rows: list[dict]) -> dict[tuple[str, str], str]:
    """Map: (date_iso, shift) -> note"""
    m: dict[tuple[str, str], str] = {}
    for r in rows:
        try:
            d = ustore.parse_iso_date(r.get("date", ""))
        except Exception:
            continue
        sh = ustore.norm_shift(r.get("shift", ""))
        if not sh:
            continue
        m[(d.isoformat(), sh)] = str(r.get("note") or "")
    return m

def _entries_map_from_tuples(entries: list[tuple[date, str, str]]) -> dict[tuple[str, str], str]:
    """Map: (date_iso, shift) -> note"""
    m: dict[tuple[str, str], str] = {}
    for d, sh, note in entries:
        if not isinstance(d, date):
            continue
        sh2 = ustore.norm_shift(sh)
        if not sh2:
            continue
        m[(d.isoformat(), sh2)] = str(note or "")
    return m

def compute_unavailability_diff(existing_rows: list[dict], new_entries: list[tuple[date, str, str]]) -> dict:
    """Return a diff summary between current store rows and the edited entries."""
    before = _entries_map_from_store_rows(existing_rows)
    after = _entries_map_from_tuples(new_entries)

    before_keys = set(before.keys())
    after_keys = set(after.keys())

    added = sorted(after_keys - before_keys)
    removed = sorted(before_keys - after_keys)

    note_changed = []
    for k in sorted(before_keys & after_keys):
        if str(before.get(k, "")) != str(after.get(k, "")):
            note_changed.append(k)

    # Privacy-friendly details: we log only (date, shift), not free-text notes.
    details = {
        "added": [{"date": k[0], "shift": k[1]} for k in added],
        "removed": [{"date": k[0], "shift": k[1]} for k in removed],
        "note_changed": [{"date": k[0], "shift": k[1]} for k in note_changed],
    }

    return {
        "before_count": len(before_keys),
        "after_count": len(after_keys),
        "added_count": len(added),
        "removed_count": len(removed),
        "note_changed_count": len(note_changed),
        "details": details,
    }

def append_unavailability_audit_log(mk: str, row: dict, max_retries: int = 3):
    """Append a row to the monthly audit log on GitHub.

    Uses optimistic concurrency (sha) and retries on conflicts.
    """
    g = _github_cfg()
    path = _audit_path_for_month(mk)

    last_err = None
    for _attempt in range(max_retries):
        gf = github_utils.get_file(
            owner=g["owner"],
            repo=g["repo"],
            path=path,
            token=g["token"],
            branch=g.get("branch", "main"),
        )
        existing_text = gf.text if gf else ""
        sha = gf.sha if gf else None

        # Build row line with proper CSV quoting
        row_buf = io.StringIO()
        writer = csv.DictWriter(row_buf, fieldnames=AUDIT_FIELDS)
        writer.writerow({k: row.get(k, "") for k in AUDIT_FIELDS})
        row_line = row_buf.getvalue().strip("\r\n")

        if existing_text.strip():
            # If header is missing/unknown, rebuild from scratch
            first = existing_text.splitlines()[0].strip()
            expected_header = ",".join(AUDIT_FIELDS)
            if first != expected_header:
                buf = io.StringIO()
                w = csv.DictWriter(buf, fieldnames=AUDIT_FIELDS)
                w.writeheader()
                # Salta il vecchio header (prima riga), riscrivi solo le righe dati
                body_lines = existing_text.splitlines()[1:]
                for line in body_lines:
                    if line.strip():
                        buf.write(line + "\n")
                buf.write(row_line + "\n")
                new_text = buf.getvalue()
            else:
                new_text = existing_text.rstrip("\n") + "\n" + row_line + "\n"
        else:
            buf = io.StringIO()
            w = csv.DictWriter(buf, fieldnames=AUDIT_FIELDS)
            w.writeheader()
            buf.write(row_line + "\n")
            new_text = buf.getvalue()

        try:
            github_utils.put_file(
                owner=g["owner"],
                repo=g["repo"],
                path=path,
                token=g["token"],
                branch=g.get("branch", "main"),
                sha=sha,
                message=f"Audit unavailability {mk}: {row.get('doctor','')}",
                text=new_text,
            )
            return  # OK
        except requests.HTTPError as e:
            last_err = e
            # retry on sha mismatch / conflict
            resp = getattr(e, "response", None)
            if _is_sha_conflict_error(e):
                continue
            raise
        except Exception as e:
            last_err = e
            raise

    if last_err:
        raise last_err

def extract_entries_from_editor(edited_rows: list[dict], yy: int, mm: int) -> tuple[list[tuple[date, str, str]], dict]:
    """Normalize and validate editor rows for a specific (yy,mm).

    Returns (entries, info) where entries is a list of (date, shift, note),
    de-duplicated by (date, shift).
    """
    entries: list[tuple[date, str, str]] = []
    invalid_date = 0
    out_of_month = 0

    for r in edited_rows or []:
        d = r.get("Data")
        if isinstance(d, datetime):
            d = d.date()
        if not isinstance(d, date):
            invalid_date += 1
            continue
        if d.year != int(yy) or d.month != int(mm):
            out_of_month += 1
            continue

        sh_raw = r.get("Fascia", "")
        sh, _changed, _unknown = normalize_fascia(sh_raw)
        sh = (sh or "").strip()
        if not sh:
            continue

        note = str(r.get("Note", "") or "")
        entries.append((d, sh, note))

    # de-duplicate by (date, shift): keep last note
    dedup: dict[tuple[date, str], str] = {}
    for d, sh, note in entries:
        dedup[(d, sh)] = note
    entries2 = [(d, sh, note) for (d, sh), note in dedup.items()]

    counts = {}
    sat_days: set[date] = set()
    sun_days: set[date] = set()
    for _d, sh, _n in entries2:
        counts[sh] = counts.get(sh, 0) + 1
        if _d.weekday() == 5:  # sabato
            sat_days.add(_d)
        elif _d.weekday() == 6:  # domenica
            sun_days.add(_d)

    return entries2, {
        "invalid_date": invalid_date,
        "out_of_month": out_of_month,
        "counts": counts,
        "sat_days": sat_days,
        "sun_days": sun_days,
    }


# ---------------- Medico UX: baseline snapshot + session guard ----------------
_BASELINE_SS_KEY = "unav_store_baseline"


def get_or_load_doctor_baseline(
    doctor: str,
    selected_months: list[tuple[int, int]],
    force_reload: bool = False,
) -> dict:
    """Return a stable baseline snapshot for the current editing session.

    The baseline is anchored in st.session_state, so we can reliably detect
    stale edits (same doctor/month saved elsewhere) at save-time.
    """
    doctor = (doctor or "").strip()
    selected_key = tuple((int(y), int(m)) for (y, m) in (selected_months or []))

    cur = st.session_state.get(_BASELINE_SS_KEY)
    if (
        (not force_reload)
        and isinstance(cur, dict)
        and cur.get("doctor") == doctor
        and tuple(cur.get("selected") or ()) == selected_key
        and isinstance(cur.get("expected_signatures"), dict)
    ):
        return cur

    # Load only this doctor's file — per-doctor storage, no cross-doctor SHA conflict.
    rows, sha = load_doctor_unavail_from_github(doctor)
    expected = _build_expected_signatures(rows, doctor, list(selected_key))
    new = {
        "doctor": doctor,
        "selected": selected_key,
        "rows": rows,
        "sha": sha,
        "expected_signatures": expected,
        "loaded_at": _utc_now_iso(),
    }
    st.session_state[_BASELINE_SS_KEY] = new
    return new


def clear_doctor_baseline():
    st.session_state.pop(_BASELINE_SS_KEY, None)


def _logout_doctor(reason: str):
    # Keep editor keys in session_state (draft), but require re-login.
    st.session_state["doctor_auth_ok"] = False
    st.session_state["doctor_name"] = None
    st.session_state["doctor_logout_msg"] = reason
    st.rerun()
    st.stop()


def _doctor_session_state_key(doctor: str) -> str:
    return f"doctor_session::{_doctor_slug(doctor)}"


def ensure_doctor_session_active(doctor: str) -> str:
    """Single-session guard per doctor.

    - On first entry: acquires/overwrites the GitHub lease (kicking out other sessions)
    - On subsequent reruns: throttled check for lease mismatch → forced logout
    - Heartbeat: periodically refreshes last_seen/expires_at on GitHub
    """
    doctor = (doctor or "").strip()
    ss_key = _doctor_session_state_key(doctor)
    now_ts = time.time()

    cur = st.session_state.get(ss_key)
    if not isinstance(cur, dict):
        cur = {
            "session_id": str(uuid.uuid4()),
            "lease_acquired": False,
            "last_check_ts": 0.0,
            "last_heartbeat_ts": 0.0,
        }
        st.session_state[ss_key] = cur

    session_id = str(cur.get("session_id") or "")
    if not session_id:
        session_id = str(uuid.uuid4())
        cur["session_id"] = session_id

    # Acquire once (overwrite existing lease) → this kicks out any other session.
    if not bool(cur.get("lease_acquired")):
        acquire_doctor_session_lease(doctor=doctor, session_id=session_id)
        cur["lease_acquired"] = True
        cur["last_check_ts"] = now_ts
        cur["last_heartbeat_ts"] = now_ts
        st.session_state[ss_key] = cur
        return session_id

    # Throttled check (avoid spamming GitHub on every data_editor rerun).
    if (now_ts - float(cur.get("last_check_ts") or 0.0)) >= DOCTOR_SESSION_CHECK_SECONDS:
        cur["last_check_ts"] = now_ts
        st.session_state[ss_key] = cur
        ok = check_doctor_session_lease(doctor, session_id)
        if not ok:
            _logout_doctor(
                "Sessione terminata: hai effettuato accesso dallo stesso utente su un altro dispositivo/browser."
            )
            st.stop()

    # Heartbeat to keep the lease alive (and also detects network/token issues).
    if (now_ts - float(cur.get("last_heartbeat_ts") or 0.0)) >= DOCTOR_SESSION_HEARTBEAT_SECONDS:
        acquire_doctor_session_lease(doctor=doctor, session_id=session_id)
        cur["last_heartbeat_ts"] = now_ts
        st.session_state[ss_key] = cur

    return session_id


def release_doctor_session(doctor: str):
    """Best-effort: mark the lease as expired when the doctor logs out."""
    doctor = (doctor or "").strip()
    ss_key = _doctor_session_state_key(doctor)
    cur = st.session_state.get(ss_key)
    if not isinstance(cur, dict):
        return
    session_id = str(cur.get("session_id") or "")
    if not session_id:
        return
    try:
        lease, sha = load_session_lease_from_github(doctor)
        if isinstance(lease, dict) and str(lease.get("session_id") or "") == session_id:
            lease["expires_at"] = _utc_now_iso()
            lease["last_seen"] = _utc_now_iso()
            save_session_lease_to_github(doctor, lease, sha, message=f"Release doctor session: {doctor}")
    except Exception:
        pass


# ---------------- UI: Header ----------------
st.title("UOC Cardiologia con UTIC - Turni")
st.markdown("""
<style>
@media (max-width: 640px) {
    .block-container { padding-left: 0.75rem !important; padding-right: 0.75rem !important; }
}
.stButton > button { white-space: normal !important; height: auto !important; }
</style>
""", unsafe_allow_html=True)

mode = st.sidebar.radio(
    "Sezione",
    ["📋 Le mie indisponibilità", "⚙️ Admin — Genera turni", "🔧 Admin — Configurazione"],
    index=0,
)

# Load default rules (for doctor list)
cfg_default = tg.load_rules(DEFAULT_RULES_PATH)
doctors_default = doctors_from_cfg(cfg_default)

# =====================================================================
#                        MEDICO – Indisponibilità
# =====================================================================
if mode == "📋 Le mie indisponibilità":
    st.subheader("Indisponibilità (Medico)")

    # GitHub is required for both indisponibilità storage and PIN self-service.
    try:
        gtmp = _github_cfg()
        if not (gtmp.get("token") and gtmp.get("owner") and gtmp.get("repo")):
            raise RuntimeError("GitHub config missing")
    except Exception:
        st.error("Archivio GitHub non configurato: configura github_unavailability in secrets.")
        st.stop()

    if not (_email_is_configured() or _sms_is_configured()):
        st.info("Nota: invio OTP via Email/SMS non configurato. Il recupero/inizializzazione PIN autonoma non sarà disponibile.")

    # ---- Session state (evita che l'app 'torni alla home' ad ogni modifica) ----
    if "doctor_auth_ok" not in st.session_state:
        st.session_state.doctor_auth_ok = False
        st.session_state.doctor_name = None

    # If this browser session was kicked out by a newer login elsewhere, show the reason.
    if st.session_state.get("doctor_logout_msg"):
        st.warning(str(st.session_state.pop("doctor_logout_msg")))

    if st.session_state.doctor_auth_ok:
        st.success(f"Accesso attivo: **{st.session_state.doctor_name}**")

        with st.expander("🔐 Cambia PIN", expanded=False):
            st.caption("Puoi cambiare il PIN in qualsiasi momento. Il nuovo PIN deve essere di 4 cifre.")
            with st.form("change_pin_form", clear_on_submit=False):
                cur_pin = st.text_input("PIN attuale", type="password", key="chg_pin_cur")
                new_pin = st.text_input("Nuovo PIN (4 cifre)", type="password", key="chg_pin_new")
                new_pin2 = st.text_input("Conferma nuovo PIN", type="password", key="chg_pin_new2")
                go_chg = st.form_submit_button("Aggiorna PIN", type="primary")
            if go_chg:
                doctor_nm = str(st.session_state.doctor_name or "")
                if not verify_doctor_pin(doctor_nm, cur_pin):
                    st.error("PIN attuale errato.")
                elif not re.fullmatch(r"\d{4}", str(new_pin or "")):
                    st.error("Il nuovo PIN deve essere di 4 cifre (solo numeri).")
                elif new_pin != new_pin2:
                    st.error("I due PIN non coincidono.")
                else:
                    try:
                        # Ensure still the active session before changing sensitive auth.
                        ensure_doctor_session_active(doctor_nm)
                        set_doctor_pin_with_retry(doctor_nm, new_pin, reason="change")
                        st.success("PIN aggiornato. Usa il nuovo PIN al prossimo accesso.")
                    except Exception as e:
                        st.error(str(e))

        if st.button("Esci / cambia medico"):

            old_doctor = str(st.session_state.doctor_name or "")
            try:
                release_doctor_session(old_doctor)
            except Exception:
                pass
            st.session_state.doctor_auth_ok = False
            st.session_state.doctor_name = None
            st.session_state.pop("doctor_selected_months", None)
            clear_doctor_baseline()
            # clear session guard state for safety
            try:
                st.session_state.pop(_doctor_session_state_key(old_doctor), None)
            except Exception:
                pass
            # cancella anche eventuali editor keys (non obbligatorio)
            st.rerun()

    if not st.session_state.doctor_auth_ok:
        st.markdown("### Accesso medico")

        doctor = st.selectbox("Seleziona il tuo nome", doctors_default, index=0, key="login_doctor")
        has_pin = doctor_has_pin(doctor)

        otp_state_key = f"otp_state::{doctor}"
        otp_state = st.session_state.get(otp_state_key) or {}

        # Tabs: login if PIN exists, otherwise first-setup
        tab_labels = ["Accedi"] + (["Primo accesso / Reset PIN"] if (_email_is_configured() or _sms_is_configured()) else [])
        if not has_pin:
            tab_labels = ["Primo accesso (Imposta PIN)"] if (_email_is_configured() or _sms_is_configured()) else ["Accesso (PIN non configurabile)"]

        tabs = st.tabs(tab_labels)

        def _do_login():
            with st.form("medico_login", clear_on_submit=False):
                pin = st.text_input("PIN", type="password", key="login_pin", help="Il tuo PIN personale (consigliato 4 cifre)")
                go = st.form_submit_button("Accedi", type="primary")
            if go:
                if verify_doctor_pin(doctor, pin):
                    clear_doctor_baseline()
                    try:
                        st.session_state.pop(_doctor_session_state_key(doctor), None)
                    except Exception:
                        pass
                    st.session_state.doctor_auth_ok = True
                    st.session_state.doctor_name = doctor
                    st.rerun()
                else:
                    st.error("PIN non valido. Controlla il PIN e riprova.")

        def _do_pin_setup_flow(mode_label: str):
            # Se hai appena aggiornato data/doctor_contacts.yml, la cache può mantenere i vecchi valori per ~2 minuti.

            if st.button("🔄 Ricarica contatti", key=f"reload_contacts_{mode_label}"):

                try:

                    load_doctor_contacts_from_github.clear()

                except Exception:

                    pass

                st.rerun()



            contacts_all = load_doctor_contacts_from_github()
            if not contacts_all:
                g = _github_cfg()
                path = g.get("contacts_path") or "data/doctor_contacts.yml"
                st.error(
                    "Non posso inviare il codice di verifica (contatti non caricati: file mancante/non leggibile oppure YAML non valido)."
                )
                st.caption(f"Sto cercando contatti in: {g.get('owner','')}/{g.get('repo','')}:{g.get('branch','main')}/{path}")
                return

            c = get_doctor_contact(doctor)

            email = str(c.get("email") or "").strip()
            phone = str(c.get("phone") or "").strip()

            available_channels = []
            if _email_is_configured() and email:
                available_channels.append(("Email", "email", _mask_email(email)))
            if _sms_is_configured() and phone:
                available_channels.append(("SMS", "sms", _mask_phone(phone)))


            if not available_channels:
                missing = []
                if not (email or phone):
                    missing.append("contatto non trovato per questo medico")
                if email and not _email_is_configured():
                    missing.append("SMTP non configurato")
                if phone and not _sms_is_configured():
                    missing.append("SMS non configurato")
                if not missing:
                    missing.append("canale non disponibile")

                st.error("Non posso inviare il codice di verifica (" + "; ".join(missing) + ").")
                # Diagnostica leggera (non mostra email/telefono)
                try:
                    _keys = list(load_doctor_contacts_from_github().keys())
                    if _keys and ("contatto non trovato per questo medico" in missing):
                        shown = ", ".join([str(k) for k in _keys[:30]])
                        more = " …" if len(_keys) > 30 else ""
                        st.caption(f"Contatti caricati per: {shown}{more}")
                except Exception:
                    pass

                if email:
                    st.caption(f"Email configurata per il medico: {_mask_email(email)}")
                if phone:
                    st.caption(f"Telefono configurato per il medico: {_mask_phone(phone)}")

                st.info(
                    "Soluzione: verifica che la chiave nel file data/doctor_contacts.yml coincida con il nome selezionato nel menu (o una variante equivalente: maiuscole/minuscole/spazi/punteggiatura non contano). "
                    "Poi configura SMTP (Gmail) o Twilio nei secrets."
                )
                return

            st.caption("Per motivi di sicurezza, per impostare o resettare il PIN serve un codice inviato via Email o SMS.")

            with st.form(f"otp_request_{mode_label}", clear_on_submit=False):
                labels = [f"{lab} ({masked})" for (lab, _ch, masked) in available_channels]
                channels = [ch for (_lab, ch, _masked) in available_channels]
                idx = 0
                choice = st.selectbox("Dove vuoi ricevere il codice?", list(range(len(channels))), format_func=lambda i: labels[i])
                send_btn = st.form_submit_button("Invia codice", type="primary")
            if send_btn:
                try:
                    masked = request_pin_otp(doctor, channels[int(choice)])
                    otp_state = {"sent_at": _utc_now_iso(), "dest": masked, "channel": channels[int(choice)]}
                    st.session_state[otp_state_key] = otp_state
                    st.success(f"Codice inviato a {masked}.")
                except Exception as e:
                    st.error(str(e))
                    return

            otp_state = st.session_state.get(otp_state_key) or {}
            if otp_state.get("dest"):
                st.info(f"Codice inviato a {otp_state.get('dest')}. Inseriscilo qui sotto per continuare.")
                with st.form(f"otp_verify_{mode_label}", clear_on_submit=False):
                    code = st.text_input("Codice (6 cifre)", key=f"otp_code_{mode_label}")
                    new_pin = st.text_input("Nuovo PIN (4 cifre)", type="password", key=f"new_pin_{mode_label}")
                    new_pin2 = st.text_input("Conferma nuovo PIN", type="password", key=f"new_pin2_{mode_label}")
                    ok = st.form_submit_button("Imposta PIN", type="primary")
                if ok:
                    if not re.fullmatch(r"\d{4}", str(new_pin or "")):
                        st.error("Il PIN deve essere di 4 cifre (solo numeri).")
                        return
                    if new_pin != new_pin2:
                        st.error("I due PIN non coincidono.")
                        return
                    try:
                        verify_pin_otp_and_consume(doctor, code)
                        set_doctor_pin_with_retry(doctor, new_pin, reason=mode_label)
                        st.session_state.pop(otp_state_key, None)
                        st.success("PIN aggiornato con successo. Ora puoi accedere.")
                    except Exception as e:
                        st.error(str(e))
                        return

        # Render tabs
        if has_pin:
            with tabs[0]:
                _do_login()
            if len(tabs) > 1:
                with tabs[1]:
                    st.markdown("#### Reset PIN (con codice via Email/SMS)")
                    _do_pin_setup_flow("reset")
        else:
            with tabs[0]:
                if (_email_is_configured() or _sms_is_configured()):
                    st.markdown("#### Primo accesso – Imposta il tuo PIN")
                    _do_pin_setup_flow("first_setup")
                else:
                    st.error("PIN non configurabile: manca configurazione Email/SMS nei secrets.")
                    st.info("Configura SMTP o Twilio per abilitare il primo accesso autonomo.")

        st.stop()

    doctor = st.session_state.doctor_name


    # Single active session per doctor: this prevents silent overwrites when the
    # same doctor uses multiple devices/browsers.
    try:
        _doctor_session_id = ensure_doctor_session_active(doctor)
    except Exception as e:
        st.error(f"Errore gestione sessione: {e}")
        st.stop()

    # ---- Selezione mesi da compilare (Anno + Mese separati) ----
    today = date.today()
    horizon_years = 20  # ampia finestra per evitare modifiche future
    year_options = list(range(today.year, today.year + horizon_years + 1))
    month_names = {
        1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio", 6: "Giugno",
        7: "Luglio", 8: "Agosto", 9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre",
    }


    # Default month/year: NEXT month relative to today's date (es. Febbraio -> Marzo)

    _first_of_this_month = today.replace(day=1)

    _first_of_next_month = (_first_of_this_month + timedelta(days=32)).replace(day=1)

    default_year = _first_of_next_month.year

    default_month = _first_of_next_month.month


    # Set defaults once per session (do not override user choices on rerun)

    st.session_state.setdefault("doctor_year_sel", default_year)

    st.session_state.setdefault("doctor_month_sel", default_month)

    st.session_state.setdefault("doctor_selected_months", [(default_year, default_month)])

    sel_default = st.session_state.get("doctor_selected_months") or [(default_year, default_month)]
    sel_set = set(sel_default)

    st.subheader("Mese da compilare")
    st.caption("Seleziona anno e mese, poi premi **▶ Aggiungi mese** per visualizzare il modulo di inserimento. Puoi aggiungere più mesi.")
    _ms_c1, _ms_c2 = st.columns([1, 2])
    with _ms_c1:
        yy_sel = st.selectbox("Anno", year_options, key="doctor_year_sel")
    with _ms_c2:
        mm_sel = st.selectbox(
            "Mese",
            list(range(1, 13)),
            format_func=lambda m: f"{m:02d} – {month_names.get(m, str(m))}",
            key="doctor_month_sel",
        )
    _ms_b1, _ms_b2 = st.columns([2, 2])
    with _ms_b1:
        add_month = st.button("▶ Aggiungi mese", use_container_width=True, help="Aggiunge l’anno/mese selezionato all’elenco.", type="primary")
    with _ms_b2:
        remove_month = st.button("✕ Rimuovi mese", use_container_width=True, help="Rimuove l’anno/mese selezionato dall’elenco.")

    cur = (int(yy_sel), int(mm_sel))
    if add_month:
        sel_set.add(cur)
        st.session_state.doctor_active_month = cur  # passa automaticamente al mese appena aggiunto
    if remove_month:
        sel_set.discard(cur)

    selected = sorted(sel_set)
    st.session_state.doctor_selected_months = selected

    if not selected:
        st.info("Seleziona anno e mese qui sopra e premi **Aggiungi mese ▶** per iniziare.")
        st.stop()

    # Gestione mese attivo (uno solo visualizzato per volta)
    st.session_state.setdefault("doctor_active_month", selected[0])
    if st.session_state.doctor_active_month not in selected:
        st.session_state.doctor_active_month = selected[0]
    active_month = st.session_state.doctor_active_month

    # Barra di navigazione tra i mesi aggiunti
    if len(selected) > 1:
        st.caption("Passa da un mese all'altro:")
        nav_cols = st.columns(min(len(selected), 3))
        for _ni, (_syy, _smm) in enumerate(selected):
            with nav_cols[_ni % 3]:
                _is_active = (_syy, _smm) == active_month
                if st.button(
                    f"{month_names.get(_smm, str(_smm))} {_syy}",
                    key=f"nav_{_syy}_{_smm}",
                    type="primary" if _is_active else "secondary",
                    use_container_width=True,
                ):
                    st.session_state.doctor_active_month = (_syy, _smm)
                    st.rerun()
        active_month = st.session_state.doctor_active_month

    # Stable baseline (snapshot) for this editing session.
    # This is what we compare against at save-time to detect a stale editor.
    refresh_baseline = st.button(
        "🔄 Ricarica dati",
        help="Ricarica l’archivio dal server (utile se qualcuno ha appena salvato).",
    )

    if refresh_baseline:
        # Reset baseline + local editors so the UI reflects the latest server state.
        clear_doctor_baseline()
        for (yy, mm) in selected:
            # legacy key (older UI)
            st.session_state.pop(f"unav_editor_{doctor}_{yy}_{mm}", None)
            # new robust UI keys (row-based widgets)
            _rows_prefix = f"unav_rows_{doctor}_{yy}_{mm}"
            for _k in list(st.session_state.keys()):
                if str(_k).startswith(_rows_prefix):
                    st.session_state.pop(_k, None)
        # Reset anche avail store baseline e editor avail
        st.session_state.pop(f"avail_store_baseline_{doctor}", None)
        for (yy, mm) in selected:
            st.session_state.pop(f"avail_rows_{doctor}_{yy}_{mm}", None)
        st.rerun()

    try:
        baseline = get_or_load_doctor_baseline(doctor, selected, force_reload=bool(refresh_baseline))
        store_rows = list(baseline.get("rows") or [])
        store_sha = baseline.get("sha")
        expected_signatures = dict(baseline.get("expected_signatures") or {})
    except Exception as e:
        st.error(f"Errore accesso archivio indisponibilità: {e}")
        st.stop()

    # Carica archivio preferenze (availability) da GitHub — solo il file del medico corrente
    _avail_base_key = f"avail_store_baseline_{doctor}"
    if _avail_base_key not in st.session_state or refresh_baseline:
        try:
            _avail_rows, _avail_sha = load_doctor_avail_from_github(doctor)
            st.session_state[_avail_base_key] = {"rows": _avail_rows, "sha": _avail_sha}
        except Exception:
            st.session_state[_avail_base_key] = {"rows": [], "sha": None}
    avail_store_rows: list[dict] = list((st.session_state.get(_avail_base_key) or {}).get("rows") or [])
    avail_store_sha: str | None = (st.session_state.get(_avail_base_key) or {}).get("sha")

    # Load app settings (open/closed + limits)
    try:
        app_settings, _settings_sha = load_app_settings_from_github()
    except Exception as e:
        app_settings, _settings_sha = dict(DEFAULT_SETTINGS), None
        st.warning(f"Impostazioni indisponibilità non leggibili (uso default): {e}")

    unav_open = bool(app_settings.get("unavailability_open", True))
    try:
        max_per_shift = int(app_settings.get("max_unavailability_per_shift", DEFAULT_SETTINGS["max_unavailability_per_shift"]))
    except Exception:
        max_per_shift = DEFAULT_SETTINGS["max_unavailability_per_shift"]
    if max_per_shift < 0:
        max_per_shift = 0

    # Cap personalizzato per questo medico (se impostato dall'admin)
    _doctor_caps = app_settings.get("doctor_caps") or {}
    try:
        max_per_shift_for_doctor = int(_doctor_caps.get(doctor, max_per_shift))
    except Exception:
        max_per_shift_for_doctor = max_per_shift
    if max_per_shift_for_doctor < 0:
        max_per_shift_for_doctor = 0

    try:
        max_weekend_days_cfg = int(app_settings.get("max_weekend_days", MAX_WEEKEND_DAYS))
    except Exception:
        max_weekend_days_cfg = MAX_WEEKEND_DAYS
    if max_weekend_days_cfg < 0:
        max_weekend_days_cfg = 0

    if not unav_open:
        st.warning("🔒 Inserimento indisponibilità temporaneamente **chiuso** dall'amministratore. Puoi solo visualizzare (non puoi salvare).")
    if max_per_shift_for_doctor != max_per_shift:
        st.caption(
            f"Limite per fascia al mese: **max {max_per_shift_for_doctor}** "
            "(limite personalizzato per il tuo profilo)."
        )
    else:
        st.caption(
            f"Limite per medico: **max {max_per_shift}** inserimenti per ogni fascia "
            "(Mattina/Pomeriggio/Notte/Diurno/Tutto il giorno) per ogni mese."
        )

    st.divider()

    edited_by_month = {}
    normalized_entries_by_month = {}
    violations_by_month = {}
    weekend_violations_by_month = {}
    info_by_month = {}
    avail_rows_by_month = {}
    save = False  # inizializzato prima del loop per sicurezza

    for (yy, mm) in [active_month]:  # mostra solo il mese attivo
        st.markdown("---")
        st.subheader(f"{month_names.get(mm, str(mm))} {yy}")
        with st.container():
            st.markdown("#### Indisponibilità")
            st.caption("Inserisci i giorni in cui NON puoi lavorare. Le righe vuote verranno ignorate.")
            existing = ustore.filter_doctor_month(store_rows, doctor, yy, mm)
            init = []
            conversions = []
            for r in existing:
                try:
                    d = datetime.fromisoformat(r["date"]).date()
                except Exception:
                    d = r["date"]
                raw_shift = r.get("shift", "")
                canon_shift, changed, unknown = normalize_fascia(raw_shift)
                if changed:
                    conversions.append({
                        "Data": d,
                        "Fascia_originale": raw_shift,
                        "Fascia_impostata": canon_shift,
                        "Nota": "Non riconosciuta (default applicato)" if unknown else "Normalizzata",
                    })
                init.append({"Data": d, "Fascia": canon_shift or "Tutto il giorno", "Note": r.get("note", "")})

            if conversions:
                st.warning("Abbiamo trovato alcune fasce non standard salvate in passato. Le abbiamo normalizzate automaticamente: controlla e, se necessario, modifica dal menu a tendina prima di salvare.")
                st.dataframe(conversions, use_container_width=True, hide_index=True)


            if not init:
                init = [{"Data": date(yy, mm, 1), "Fascia": "Mattina", "Note": ""}]

            # --- Editor righe (UI robusta: niente st.data_editor) ---
            rows_key = f"unav_rows_{doctor}_{yy}_{mm}"

            # Initialize per-month rows state once (or after refresh)
            if rows_key not in st.session_state:
                rows_init = []
                for _r in init:
                    rows_init.append({
                        "id": str(uuid.uuid4()),
                        "Data": _r.get("Data"),
                        "Fascia": _r.get("Fascia") or "Mattina",
                        "Note": _r.get("Note", ""),
                    })
                st.session_state[rows_key] = rows_init

            first_day = date(yy, mm, 1)
            if mm == 12:
                last_day = date(yy + 1, 1, 1) - timedelta(days=1)
            else:
                last_day = date(yy, mm + 1, 1) - timedelta(days=1)

            if unav_open:
                rows = list(st.session_state.get(rows_key) or [])

                if not rows:
                    rows = [{
                        "id": str(uuid.uuid4()),
                        "Data": first_day,
                        "Fascia": "Mattina",
                        "Note": "",
                    }]
                    st.session_state[rows_key] = rows

                with st.expander("📅 Aggiungi periodo ferie", expanded=False):
                    st.caption(
                        "Seleziona un intervallo di date e premi **Aggiungi giorni** per inserire ogni giorno "
                        "del periodo come riga con fascia *Ferie* nella tabella sottostante. "
                        "Puoi usare questo strumento più volte per aggiungere periodi separati. "
                        "I sabati e le domeniche in ferie contano nel limite di "
                        f"{max_weekend_days_cfg} sabati e {max_weekend_days_cfg} domeniche al mese."
                    )
                    _fp_col1, _fp_col2, _fp_col3 = st.columns([2, 2, 1], vertical_alignment="bottom")
                    with _fp_col1:
                        _ferie_start = st.date_input(
                            "Dal",
                            value=first_day,
                            min_value=first_day,
                            max_value=last_day,
                            key=f"{rows_key}__ferie_start",
                            format="DD/MM/YYYY",
                        )
                    with _fp_col2:
                        _ferie_end = st.date_input(
                            "Al",
                            value=first_day,
                            min_value=first_day,
                            max_value=last_day,
                            key=f"{rows_key}__ferie_end",
                            format="DD/MM/YYYY",
                        )
                    with _fp_col3:
                        _add_ferie = st.button(
                            "Aggiungi giorni",
                            key=f"{rows_key}__ferie_add",
                            use_container_width=True,
                        )
                    if _add_ferie:
                        if _ferie_end < _ferie_start:
                            st.error("La data di fine deve essere uguale o successiva alla data di inizio.")
                        else:
                            _current_rows = list(st.session_state.get(rows_key) or [])
                            _existing_ferie_dates = {
                                r["Data"] for r in _current_rows if r.get("Fascia") == "Ferie"
                            }
                            _delta = (_ferie_end - _ferie_start).days + 1
                            for _offset in range(_delta):
                                _d = _ferie_start + timedelta(days=_offset)
                                if _d not in _existing_ferie_dates:
                                    _current_rows.append({
                                        "id": str(uuid.uuid4()),
                                        "Data": _d,
                                        "Fascia": "Ferie",
                                        "Note": "",
                                    })
                            st.session_state[rows_key] = _current_rows
                            st.rerun()

                # Header
                h1, h2, h3 = st.columns([2, 2, 1])
                h1.markdown("**Data**")
                h2.markdown("**Fascia**")
                h3.markdown("**Rimuovi**")

                remove_ids = []
                new_rows = []

                for r in rows:
                    rid = str(r.get("id") or uuid.uuid4())
                    d_key = f"{rows_key}__d__{rid}"
                    s_key = f"{rows_key}__s__{rid}"
                    n_key = f"{rows_key}__n__{rid}"
                    rm_key = f"{rows_key}__rm__{rid}"

                    # Seed widget state once to avoid "first change gets lost" behaviour
                    if d_key not in st.session_state:
                        st.session_state[d_key] = r.get("Data") or first_day
                    if s_key not in st.session_state:
                        st.session_state[s_key] = r.get("Fascia") or "Mattina"
                    if n_key not in st.session_state:
                        st.session_state[n_key] = r.get("Note", "")

                    c1, c2, c3 = st.columns([2, 2, 1], vertical_alignment="bottom")
                    with c1:
                        d_val = st.date_input(
                            "Data",
                            key=d_key,
                            min_value=first_day,
                            max_value=last_day,
                            label_visibility="collapsed",
                            format="DD/MM/YYYY",
                        )
                    with c2:
                        sh_val = st.selectbox(
                            "Fascia",
                            options=FASCIA_OPTIONS,
                            key=s_key,
                            label_visibility="collapsed",
                        )
                    with c3:
                        if st.button("🗑️", key=rm_key, help="Rimuovi questa riga"):
                            remove_ids.append(rid)
                    _has_note = bool(str(st.session_state.get(n_key, "") or "").strip())
                    with st.expander("📝 Note", expanded=_has_note):
                        note_val = st.text_input(
                            "Note",
                            key=n_key,
                            label_visibility="collapsed",
                        )

                    # Enforce month bounds (extra safety)
                    if isinstance(d_val, date):
                        if d_val < first_day:
                            d_val = first_day
                            st.session_state[d_key] = d_val
                        if d_val > last_day:
                            d_val = last_day
                            st.session_state[d_key] = d_val

                    new_rows.append({
                        "id": rid,
                        "Data": d_val,
                        "Fascia": sh_val,
                        "Note": note_val,
                    })

                if remove_ids:
                    new_rows = [r for r in new_rows if str(r.get("id")) not in set(remove_ids)]
                    if not new_rows:
                        new_rows = [{
                            "id": str(uuid.uuid4()),
                            "Data": first_day,
                            "Fascia": "Mattina",
                            "Note": "",
                        }]
                    st.session_state[rows_key] = new_rows
                    st.rerun()

                # Persist updated rows for this rerun (safe: not a widget key)
                st.session_state[rows_key] = new_rows

                # Build "edited" compatible with existing save/validation pipeline
                edited = [{"Data": rr["Data"], "Fascia": rr["Fascia"], "Note": rr.get("Note", "")} for rr in new_rows]
            else:
                # Read-only view when the admin closes submissions
                st.dataframe(init, use_container_width=True, hide_index=True)
                edited = init

            edited_by_month[(yy, mm)] = edited

            # Normalize & validate + enforce max per shift (per month)
            entries_norm, info = extract_entries_from_editor(edited, yy, mm)
            normalized_entries_by_month[(yy, mm)] = entries_norm
            info_by_month[(yy, mm)] = info

            counts = info.get("counts", {}) or {}
            sat_days = info.get("sat_days", set())
            sun_days = info.get("sun_days", set())
            over = {sh: n for sh, n in counts.items() if sh != "Ferie" and n > max_per_shift_for_doctor}
            weekend_over = {}
            if len(sat_days) > max_weekend_days_cfg:
                weekend_over["Sabati"] = len(sat_days)
            if len(sun_days) > max_weekend_days_cfg:
                weekend_over["Domeniche"] = len(sun_days)
            violations_by_month[(yy, mm)] = over
            weekend_violations_by_month[(yy, mm)] = weekend_over

            if info.get("out_of_month"):
                st.warning(
                    f"⚠️ {info['out_of_month']} righe con data fuori mese sono state ignorate "
                    f"(devono essere in {yy}-{mm:02d})."
                )
            if info.get("invalid_date"):
                st.warning(f"⚠️ {info['invalid_date']} righe hanno una data non valida e sono state ignorate.")

            _fascia_str = " · ".join([
                f"{sh} {counts.get(sh, 0)}/{'∞' if sh == 'Ferie' else max_per_shift_for_doctor}"
                for sh in FASCIA_OPTIONS
                if counts.get(sh, 0) > 0
            ]) or "nessuna indisponibilità inserita"
            st.caption(f"Fasce: {_fascia_str}")
            st.caption(f"Weekend: Sabati {len(sat_days)}/{max_weekend_days_cfg} · Domeniche {len(sun_days)}/{max_weekend_days_cfg}")

            if over:
                pretty = ", ".join([
                    f"{sh}: {n}/{'∞' if sh == 'Ferie' else max_per_shift_for_doctor}"
                    for sh, n in over.items()
                ])
                st.error(f"Limite superato in questo mese → {pretty}. Rimuovi alcune righe prima di salvare.")

            if weekend_over:
                we_pretty = ", ".join([f"{label}: {n}/{max_weekend_days_cfg}" for label, n in weekend_over.items()])
                st.error(f"Limite weekend superato → {we_pretty}. Puoi segnare al massimo {max_weekend_days_cfg} sabati e {max_weekend_days_cfg} domeniche al mese.")

            # ── Pulsante salva indisponibilità (tra le due sezioni) ──────────────
            _can_save = bool(unav_open) and not bool(over) and not bool(weekend_over)
            render_unav_flash(doctor)
            _sc1, _sc2, _sc3 = st.columns([3, 2, 2])
            with _sc1:
                save = st.button(
                    "💾 Salva indisponibilità",
                    key=f"save_unav_{yy}_{mm}",
                    type="primary",
                    disabled=not _can_save,
                    use_container_width=True,
                )
            with _sc2:
                add_row = st.button("➕ Aggiungi riga", key=f"{rows_key}__add", use_container_width=True, disabled=not unav_open)
            with _sc3:
                clean_rows = st.button("🧹 Pulisci vuote", key=f"{rows_key}__clean", use_container_width=True, disabled=not unav_open)

            if add_row:
                new_rows_state = list(st.session_state.get(rows_key) or [])
                new_rows_state.append({
                    "id": str(uuid.uuid4()),
                    "Data": first_day,
                    "Fascia": "Mattina",
                    "Note": "",
                })
                st.session_state[rows_key] = new_rows_state
                st.rerun()

            if clean_rows:
                def _is_empty(_x: dict) -> bool:
                    d = _x.get("Data")
                    sh = str(_x.get("Fascia") or "").strip()
                    note = str(_x.get("Note") or "").strip()
                    # Una riga è "vuota" se ha la data di default (primo del mese) e nessuna nota
                    is_default_date = isinstance(d, date) and d.day == 1
                    return is_default_date and sh in ("Mattina", "") and not note

                _cleaned = [r for r in (st.session_state.get(rows_key) or []) if not _is_empty(r)]
                if not _cleaned:
                    _cleaned = [{
                        "id": str(uuid.uuid4()),
                        "Data": first_day,
                        "Fascia": "Mattina",
                        "Note": "",
                    }]
                st.session_state[rows_key] = _cleaned
                st.rerun()

            # ── Disponibilità (preferenze) ──────────────────────────────────────────
            st.divider()
            st.markdown("#### Disponibilità (preferenze)")
            with st.container():
                max_avail = int(app_settings.get("max_availability_per_shift", 6))
                st.caption(
                    f"Inserisci i giorni/fasce in cui **preferiresti** lavorare. "
                    f"Il software **proverà** (senza garanzia) a rispettarle. "
                    f"Limite: max **{max_avail}** per fascia al mese."
                )
                avail_key = f"avail_rows_{doctor}_{yy}_{mm}"
                if avail_key not in st.session_state:
                    _existing_avail = ustore.filter_doctor_month(avail_store_rows, doctor, yy, mm)
                    if _existing_avail:
                        st.session_state[avail_key] = [
                            {
                                "id": str(uuid.uuid4()),
                                "Data": date.fromisoformat(r["date"]) if r.get("date") else date(yy, mm, 1),
                                "Fascia": r.get("shift", "Mattina"),
                                "Priorita": r.get("priority", "media"),
                                "Note": r.get("note", ""),
                            }
                            for r in _existing_avail
                        ]
                    else:
                        st.session_state[avail_key] = [
                            {"id": str(uuid.uuid4()), "Data": date(yy, mm, 1), "Fascia": "Mattina",
                             "Priorita": "media", "Note": ""}
                        ]

                if unav_open:
                    _PRIORITY_OPTIONS = ["media", "alta", "bassa"]
                    _PRIORITY_LABELS = {"alta": "⬆ Alta", "media": "● Media", "bassa": "⬇ Bassa"}
                    av_rows = list(st.session_state.get(avail_key) or [])
                    updated_av = []
                    for av_r in av_rows:
                        # Riga 1: Data, Fascia, bottone elimina
                        _av_r1c1, _av_r1c2, _av_r1c3 = st.columns([2, 2, 0.5], vertical_alignment="bottom")
                        with _av_r1c1:
                            av_date = st.date_input(
                                "Data",
                                value=av_r.get("Data") or date(yy, mm, 1),
                                min_value=date(yy, mm, 1),
                                max_value=date(yy + 1, 1, 1) - timedelta(days=1) if mm == 12 else date(yy, mm + 1, 1) - timedelta(days=1),
                                key=f"{avail_key}_{av_r['id']}_d",
                                format="DD/MM/YYYY",
                            )
                        with _av_r1c2:
                            av_shift = st.selectbox(
                                "Fascia",
                                AVAIL_FASCIA_OPTIONS,
                                index=AVAIL_FASCIA_OPTIONS.index(av_r.get("Fascia", "Mattina"))
                                      if av_r.get("Fascia", "Mattina") in AVAIL_FASCIA_OPTIONS else 0,
                                key=f"{avail_key}_{av_r['id']}_s",
                            )
                        with _av_r1c3:
                            del_av = st.button("🗑", key=f"{avail_key}_{av_r['id']}_del")

                        # Riga 2 (collapsibile): Priorità e Note
                        _prev_pri = av_r.get("Priorita", "media")
                        _pri_idx = _PRIORITY_OPTIONS.index(_prev_pri) if _prev_pri in _PRIORITY_OPTIONS else 0
                        _has_extra = _prev_pri != "media" or bool(av_r.get("Note", "").strip())
                        with st.expander("Priorità / Note", expanded=_has_extra):
                            _av_r2c1, _av_r2c2 = st.columns([1, 2])
                            with _av_r2c1:
                                av_priority = st.selectbox(
                                    "Priorità",
                                    [_PRIORITY_LABELS[p] for p in _PRIORITY_OPTIONS],
                                    index=_pri_idx,
                                    key=f"{avail_key}_{av_r['id']}_p",
                                )
                                av_priority_val = _PRIORITY_OPTIONS[
                                    [_PRIORITY_LABELS[p] for p in _PRIORITY_OPTIONS].index(av_priority)
                                ]
                            with _av_r2c2:
                                av_note = st.text_input(
                                    "Note",
                                    value=av_r.get("Note", ""),
                                    key=f"{avail_key}_{av_r['id']}_n",
                                )

                        if not del_av:
                            updated_av.append({
                                "id": av_r["id"],
                                "Data": av_date,
                                "Fascia": av_shift,
                                "Priorita": av_priority_val,
                                "Note": av_note,
                            })
                    st.session_state[avail_key] = updated_av

                    # Conta per fascia
                    av_counts = {}
                    for r in (st.session_state.get(avail_key) or []):
                        sh = r.get("Fascia","")
                        if sh: av_counts[sh] = av_counts.get(sh, 0) + 1
                    av_over = {sh: n for sh, n in av_counts.items() if n > max_avail}
                    if av_over:
                        st.error(f"Limite disponibilità superato: {av_over}. Rimuovi alcune righe.")
                    else:
                        st.caption("Conteggi: " + ", ".join([f"{sh} {av_counts.get(sh,0)}/{max_avail}" for sh in AVAIL_FASCIA_OPTIONS if av_counts.get(sh,0)>0]))

                    st.divider()
                    _save_avail_disabled = bool(av_over)
                    _av_sc1, _av_sc2, _av_sc3 = st.columns([3, 2, 2])
                    with _av_sc1:
                        _do_save_avail = st.button("💾 Salva preferenze", key=f"save_avail_{doctor}_{yy}_{mm}", type="primary", disabled=_save_avail_disabled, use_container_width=True)
                    with _av_sc2:
                        if st.button("➕ Aggiungi", key=f"{avail_key}__add", use_container_width=True):
                            st.session_state[avail_key].append({
                                "id": str(uuid.uuid4()),
                                "Data": date(yy, mm, 1),
                                "Fascia": "Mattina",
                                "Note": "",
                            })
                            st.rerun()
                    with _av_sc3:
                        if st.button("🧹 Pulisci", key=f"{avail_key}__clean", use_container_width=True):
                            st.session_state[avail_key] = [
                                r for r in st.session_state[avail_key]
                                if r.get("Data") or str(r.get("Note","")).strip()
                            ] or [{"id": str(uuid.uuid4()), "Data": date(yy, mm, 1), "Fascia": "Mattina", "Note": ""}]
                            st.rerun()
                    if _do_save_avail:
                        _avail_entries = [
                            (r["Data"], r.get("Fascia", "Mattina"), r.get("Note", ""), r.get("Priorita", "media"))
                            for r in (st.session_state.get(avail_key) or [])
                            if r.get("Data") and r.get("Fascia")
                        ]
                        _avail_upd = datetime.utcnow().isoformat(timespec="seconds") + "Z"
                        try:
                            _new_avail_sha = save_doctor_availability_with_retry(
                                doctor=doctor,
                                entries_by_month={(yy, mm): _avail_entries},
                                updated_at=_avail_upd,
                                message=f"Update availability: {doctor} ({_avail_upd})",
                                initial_rows=avail_store_rows,
                                initial_sha=avail_store_sha,
                            )
                            # Ricarica il file del medico per aggiornare SHA e rows
                            _avail_fresh_rows, _avail_fresh_sha = load_doctor_avail_from_github(doctor)
                            st.session_state[f"avail_store_baseline_{doctor}"] = {
                                "rows": _avail_fresh_rows,
                                "sha": _avail_fresh_sha or _new_avail_sha,
                            }
                            st.success(f"Preferenze salvate ({len(_avail_entries)} voci).")
                        except Exception as _e:
                            st.error(f"Errore salvataggio preferenze: {_e}")

                    # Salva in sessione per trasmetterle al generate (chiave globale con medico)
                    avail_rows_by_month[(yy, mm)] = [
                        {"date": str(r["Data"]), "shift": r["Fascia"], "priority": r.get("Priorita", "media")}
                        for r in (st.session_state.get(avail_key) or [])
                        if r.get("Data") and r.get("Fascia")
                    ]
                    # Persiste nella sessione globale per l'admin
                    global_avail_key = f"avail_global_{doctor}_{yy}_{mm}"
                    st.session_state[global_avail_key] = [
                        {"doctor": doctor, "date": str(r["Data"]), "shift": r["Fascia"],
                         "priority": r.get("Priorita", "media")}
                        for r in (st.session_state.get(avail_key) or [])
                        if r.get("Data") and r.get("Fascia")
                    ]
                else:
                    # Read-only view when the admin closes submissions
                    _existing_ro = ustore.filter_doctor_month(avail_store_rows, doctor, yy, mm)
                    if _existing_ro:
                        _ro_df = [
                            {
                                "Data": r.get("date", ""),
                                "Fascia": r.get("shift", ""),
                                "Note": r.get("note", ""),
                            }
                            for r in _existing_ro
                        ]
                        st.info("🔒 Inserimento chiuso dall'amministratore. Visualizzazione in sola lettura.")
                        st.dataframe(_ro_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("🔒 Inserimento chiuso dall'amministratore. Nessuna preferenza salvata per questo mese.")

    if save:
        if not unav_open:
            st.error("Inserimento indisponibilità chiuso dall'amministratore: non è possibile salvare.")
            st.stop()

        # Force a lease ownership check right before saving (no throttle).
        try:
            _ss_key = _doctor_session_state_key(doctor)
            _cur = st.session_state.get(_ss_key) if isinstance(st.session_state.get(_ss_key), dict) else {}
            _sid = str((_cur or {}).get("session_id") or "")
            if not _sid:
                # Should not happen, but be safe.
                _sid = ensure_doctor_session_active(doctor)
            if not check_doctor_session_lease(doctor, _sid):
                _logout_doctor(
                    "Impossibile salvare: la sessione è stata sostituita da un accesso dello stesso utente da un altro dispositivo/browser."
                )
                st.stop()
            # refresh lease timestamp so an in-progress save won't appear "expired"
            acquire_doctor_session_lease(doctor=doctor, session_id=_sid)
        except Exception as e:
            st.error(f"Errore verifica sessione prima del salvataggio: {e}")
            st.stop()

        # Server-side re-check (in caso di race / rerun)
        hard_viol = []
        for (yy, mm), entries_norm in (normalized_entries_by_month or {}).items():
            counts = {}
            sat_days_s: set[date] = set()
            sun_days_s: set[date] = set()
            for _d, sh, _n in entries_norm:
                counts[sh] = counts.get(sh, 0) + 1
                if _d.weekday() == 5:
                    sat_days_s.add(_d)
                elif _d.weekday() == 6:
                    sun_days_s.add(_d)
            over = {sh: n for sh, n in counts.items() if sh != "Ferie" and n > max_per_shift_for_doctor}
            if over:
                hard_viol.append(
                    f"{yy}-{mm:02d}: " + ", ".join([f"{sh} {n}/{'∞' if sh == 'Ferie' else max_per_shift_for_doctor}" for sh, n in over.items()])
                )
            if len(sat_days_s) > max_weekend_days_cfg:
                hard_viol.append(f"{yy}-{mm:02d}: Sabati {len(sat_days_s)}/{max_weekend_days_cfg}")
            if len(sun_days_s) > max_weekend_days_cfg:
                hard_viol.append(f"{yy}-{mm:02d}: Domeniche {len(sun_days_s)}/{max_weekend_days_cfg}")

        if hard_viol:
            st.error(
                "Impossibile salvare: limite indisponibilità superato.\n\n"
                + "\n".join([f"- {x}" for x in hard_viol])
            )
            st.stop()

        updated_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"

        try:
            audit_todo, _final_sha = save_doctor_unavailability_with_retry(
                doctor=doctor,
                normalized_entries_by_month=normalized_entries_by_month or {},
                updated_at=updated_at,
                message=f"Update unavailability: {doctor} ({updated_at})",
                initial_rows=store_rows,
                initial_sha=store_sha,
                expected_signatures=expected_signatures,
                max_retries=6,
            )

            # Monthly audit log (best-effort)
            for mk_audit, diff in audit_todo:
                audit_row = {
                    "ts_utc": updated_at,
                    "doctor": doctor,
                    "month": mk_audit,
                    "action": "save",
                    "before_count": diff.get("before_count", 0),
                    "after_count": diff.get("after_count", 0),
                    "added_count": diff.get("added_count", 0),
                    "removed_count": diff.get("removed_count", 0),
                    "note_changed_count": diff.get("note_changed_count", 0),
                    "details_json": json.dumps(diff.get("details", {}), ensure_ascii=False),
                    "app_build": APP_BUILD,
                }
                try:
                    append_unavailability_audit_log(mk_audit, audit_row)
                except Exception as e:
                    st.warning(f"Audit log non aggiornato per {mk_audit}: {e}")

            # After a successful save, refresh our baseline snapshot so next saves
            # don't trigger false "stale" conflicts.
            clear_doctor_baseline()

            # Build informative success message
            total_added = sum(d.get("added_count", 0) for _, d in audit_todo)
            total_removed = sum(d.get("removed_count", 0) for _, d in audit_todo)
            months_changed = [mk for mk, _ in audit_todo]
            if months_changed:
                parts = []
                if total_added:
                    parts.append(f"+{total_added} aggiunt{'a' if total_added == 1 else 'e'}")
                if total_removed:
                    parts.append(f"-{total_removed} rimoss{'a' if total_removed == 1 else 'e'}")
                detail = f" ({', '.join(parts)})" if parts else ""
                mesi_str = ", ".join(months_changed)
                success_msg = f"Salvataggio effettuato ✅ — mesi aggiornati: {mesi_str}{detail}"
            else:
                success_msg = "Salvataggio effettuato ✅ — nessuna modifica rispetto ai dati già presenti"
            set_unav_flash(doctor, "success", success_msg)
            st.rerun()
        except Exception as e:
            set_unav_flash(
                doctor,
                "error",
                f"Errore durante il salvataggio ❌ — {type(e).__name__}: {e}",
                details=(
                    "Se il problema persiste: ricarica la pagina e riprova.\n\n"
                    "Se vedi 404: (1) token senza accesso alla repo privata, "
                    "(2) owner/repo/branch/path errati, "
                    "(3) token non autorizzato SSO (se repo in Organization)."
                ),
            )
            st.rerun()


# =====================================================================
#                           ADMIN
# =====================================================================
else:
    st.subheader("Area Admin")
    admin_pin = _get_admin_pin()
    if not admin_pin:
        st.error("Admin PIN non configurato in secrets (auth.admin_pin).")
        st.stop()

    # Persist admin auth across reruns
    if "admin_auth_ok" not in st.session_state:
        st.session_state.admin_auth_ok = False

    if not st.session_state.admin_auth_ok:
        with st.form("admin_login"):
            pin = st.text_input("PIN Admin", type="password")
            ok = st.form_submit_button("Sblocca area Admin", type="primary")

        if not ok:
            st.stop()
        if pin != admin_pin:
            st.error("PIN Admin errato.")
            st.stop()

        st.session_state.admin_auth_ok = True
        # Rerun to avoid re-submitting the form on next widget interaction
        st.rerun()

    col_logout, col_status = st.columns([1, 3])
    with col_logout:
        if st.button("Esci (Admin)", help="Chiude la sessione Admin su questo browser."):
            st.session_state.admin_auth_ok = False
            st.rerun()
    with col_status:
        st.success("Area Admin sbloccata ✅")

    # Carica cfg una volta per entrambi i rami (Genera e Configurazione)
    cfg_admin = tg.load_rules(DEFAULT_RULES_PATH)
    doctors = doctors_from_cfg(cfg_admin)
    rules_path = DEFAULT_RULES_PATH

    # ── Ramo Configurazione ───────────────────────────────────────────────
    if mode == "🔧 Admin — Configurazione":
        st.markdown("### 🔧 Configurazione sistema")

        with st.expander("⚙️ Impostazioni indisponibilità", expanded=False):
            try:
                app_settings, app_settings_sha = load_app_settings_from_github()
            except Exception as e:
                app_settings, app_settings_sha = dict(DEFAULT_SETTINGS), None
                st.warning(f"Impossibile leggere impostazioni da GitHub (uso default): {e}")

            cur_open = bool(app_settings.get("unavailability_open", True))
            try:
                cur_max = int(app_settings.get("max_unavailability_per_shift", DEFAULT_SETTINGS["max_unavailability_per_shift"]))
            except Exception:
                cur_max = DEFAULT_SETTINGS["max_unavailability_per_shift"]
            if cur_max < 0:
                cur_max = 0

            new_open = st.toggle(
                "Consenti ai medici di inserire/modificare indisponibilità",
                value=cur_open,
                help="Se disattivato, i medici possono solo visualizzare le proprie indisponibilità ma non salvarle.",
            )
            _aS1, _aS2, _aS3, _aS4 = st.columns([1, 1, 1, 1.5])
            with _aS1:
                new_max = st.number_input(
                    "Max indisponibilità/fascia",
                    min_value=0,
                    max_value=31,
                    value=int(cur_max),
                    step=1,
                    help="Esempio: 6 = max 6 Mattine, 6 Pomeriggi, ecc. per ogni mese.",
                )
            with _aS2:
                try:
                    cur_max_avail = int(app_settings.get("max_availability_per_shift", DEFAULT_SETTINGS["max_availability_per_shift"]))
                except Exception:
                    cur_max_avail = DEFAULT_SETTINGS["max_availability_per_shift"]
                new_max_avail = st.number_input(
                    "Max disponibilità/fascia",
                    min_value=0,
                    max_value=31,
                    value=int(cur_max_avail),
                    step=1,
                    help="Max preferenze 'disponibilità' inseribili per fascia per mese.",
                )
            with _aS3:
                try:
                    cur_max_weekend = int(app_settings.get("max_weekend_days", DEFAULT_SETTINGS["max_weekend_days"]))
                except Exception:
                    cur_max_weekend = DEFAULT_SETTINGS["max_weekend_days"]
                new_max_weekend = st.number_input(
                    "Max weekend/mese",
                    min_value=0,
                    max_value=5,
                    value=int(cur_max_weekend),
                    step=1,
                    help="Max sabati distinti e max domeniche distinte che ogni medico può segnare come indisponibile in un mese (Ferie incluse).",
                )
            with _aS4:
                meta = ""
                if app_settings.get("updated_at"):
                    meta += f"Ultimo aggiornamento: {app_settings.get('updated_at')}"
                if app_settings.get("updated_by"):
                    meta += f" | da: {app_settings.get('updated_by')}"
                if meta:
                    st.caption(meta)

            # Cap personalizzati per universitari
            st.markdown("**Cap personalizzati per universitari** *(sovrascrivono il limite globale per i medici selezionati)*")
            gc_uni = (cfg_admin.get("global_constraints") or {}).get("university_doctors") or {}
            uni_doctors = sorted(gc_uni.keys()) if gc_uni else ["Dattilo", "De Gregorio", "Zito"]
            cur_doctor_caps = app_settings.get("doctor_caps") or {}
            new_doctor_caps = {}
            uni_cols = st.columns(len(uni_doctors))
            for col, doc in zip(uni_cols, uni_doctors):
                with col:
                    cur_cap = cur_doctor_caps.get(doc, int(new_max))
                    new_doctor_caps[doc] = st.number_input(
                        doc,
                        min_value=0,
                        max_value=31,
                        value=int(cur_cap),
                        step=1,
                        key=f"doctor_cap_{doc}",
                        help=f"Cap massimo di indisponibilità per fascia al mese per {doc}. Usa il limite globale ({int(new_max)}) se non vuoi differenziare.",
                    )

            if st.button("Salva impostazioni indisponibilità", type="primary"):
                settings_to_save = {
                    "unavailability_open": bool(new_open),
                    "max_unavailability_per_shift": int(new_max),
                    "max_availability_per_shift": int(new_max_avail),
                    "max_weekend_days": int(new_max_weekend),
                    "doctor_caps": {doc: int(v) for doc, v in new_doctor_caps.items()},
                    "updated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    "updated_by": "admin",
                }
                try:
                    save_app_settings_to_github(
                        settings_to_save,
                        app_settings_sha,
                        message=f"Update settings: open={bool(new_open)} max_unav={int(new_max)} max_avail={int(new_max_avail)} max_weekend={int(new_max_weekend)}",
                    )
                    st.success("Impostazioni salvate ✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore salvataggio impostazioni su GitHub: {e}")

        # ── Gestione Pool Medici ──────────────────────────────────────────
        with st.expander("🩺 Gestione Pool Medici", expanded=False):
            _pool_cfg_loaded, _pool_cfg_sha = load_pool_config_from_github_st()
            _pool_cfg_exists = bool(_pool_cfg_loaded)

            if not _pool_cfg_exists:
                st.warning("Nessuna configurazione pool trovata su GitHub. Inizializza dal YAML attuale per cominciare.", icon="⚠️")
                if st.button("🔧 Inizializza da YAML attuale", key="btn_init_pool_cfg"):
                    import pool_config_store as _pcs_init
                    _migrated = _pcs_init.migrate_from_yaml(cfg_admin)
                    _ok, _msg = save_pool_config_with_retry(_migrated, None)
                    if _ok:
                        st.success("Configurazione pool inizializzata dal YAML. Ricarica la pagina per modificarla.", icon="✅")
                        st.rerun()
                    else:
                        st.error(_msg)
            else:
                _pool_draft_key = "pool_cfg_draft"
                if _pool_draft_key not in st.session_state:
                    import copy as _copy_pool
                    st.session_state[_pool_draft_key] = _copy_pool.deepcopy(_pool_cfg_loaded)

                _draft = st.session_state[_pool_draft_key]
                _draft_doctors: dict = _draft.get("doctors", {})
                _all_cols = sorted(cfg_admin.get("columns", {}).keys())
                _all_docs_list = sorted(_draft_doctors.keys(), key=lambda s: (s == "Recupero", s.lower()))

                import pool_config_store as _pcs_ui
                _tab_med, _tab_col, _tab_lim, _tab_serv = st.tabs(
                    ["👨‍⚕️ Medici", "📋 Colonne", "⚖️ Limiti", "🔗 Servizi"]
                )

                with _tab_med:
                    st.markdown("**Stato e flag per ogni medico** — aggiungi righe con ＋, elimina con ✕")
                    import pandas as _pd_pool
                    _med_rows = []
                    for _dname in _all_docs_list:
                        _dc = _draft_doctors[_dname]
                        _med_rows.append({
                            "Medico": _dname,
                            "Attivo": bool(_dc.get("active", True)),
                            "Reperibilità": not bool(_dc.get("excluded_from_reperibilita", False)),
                            "Festivi diurni": bool(_dc.get("festivi_diurni", True)),
                            "Festivi notti": bool(_dc.get("festivi_notti", True)),
                            "Universitario": bool(_dc.get("university_doctor")),
                            "Ratio uni": float((_dc.get("university_doctor") or {}).get("ratio", 0.6) or 0.6)
                            if _dc.get("university_doctor") else 0.0,
                        })
                    _med_df = _pd_pool.DataFrame(_med_rows)
                    _edited_med = st.data_editor(
                        _med_df,
                        column_config={
                            "Medico": st.column_config.TextColumn("Medico", help="Cognome esatto (maiuscola iniziale)"),
                            "Attivo": st.column_config.CheckboxColumn("Attivo"),
                            "Reperibilità": st.column_config.CheckboxColumn("Reperibilità C"),
                            "Festivi diurni": st.column_config.CheckboxColumn("Festivi diurni"),
                            "Festivi notti": st.column_config.CheckboxColumn("Festivi notti"),
                            "Universitario": st.column_config.CheckboxColumn("Universitario"),
                            "Ratio uni": st.column_config.NumberColumn("Ratio", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"),
                        },
                        hide_index=True,
                        use_container_width=True,
                        key="pool_med_editor",
                        num_rows="dynamic",
                    )
                    _edited_names: set[str] = set()
                    for _, _row in _edited_med.iterrows():
                        _dn = (_row.get("Medico") or "").strip()
                        if not _dn:
                            continue
                        _edited_names.add(_dn)
                        if _dn not in _draft_doctors:
                            _draft_doctors[_dn] = {
                                "active": True, "columns": [],
                                "festivi_diurni": True, "festivi_notti": True,
                                "excluded_from_reperibilita": False,
                                "university_doctor": None, "column_overrides": {},
                            }
                        _draft_doctors[_dn]["active"] = bool(_row["Attivo"])
                        _draft_doctors[_dn]["excluded_from_reperibilita"] = not bool(_row["Reperibilità"])
                        _draft_doctors[_dn]["festivi_diurni"] = bool(_row["Festivi diurni"])
                        _draft_doctors[_dn]["festivi_notti"] = bool(_row["Festivi notti"])
                        _is_uni = bool(_row["Universitario"])
                        if _is_uni:
                            _draft_doctors[_dn]["university_doctor"] = {"ratio": float(_row["Ratio uni"] or 0.6)}
                        else:
                            _draft_doctors[_dn]["university_doctor"] = None
                    for _dn_old in list(_draft_doctors.keys()):
                        if _dn_old not in _edited_names:
                            del _draft_doctors[_dn_old]

                    st.divider()
                    with st.expander("📌 Vincoli strutturali fissi (sola lettura — modificabili solo da YAML)", expanded=False):
                        st.info(
                            "Questi vincoli sono hardcoded nel YAML e non modificabili da questa GUI:\n\n"
                            "- **Cimino** esatto 2 turni U al mese\n"
                            "- **Crea** unico medico per i sabati AB (2/mese)\n"
                            "- **Allegra** vincolo lunedì V+U (stessa giornata)\n"
                            "- **De Gregorio** max 3 giorni feriali su I\n"
                            "- **Grimaldi e Calabrò** esenti dal vincolo 'min 2 weekend liberi/mese'\n"
                            "- **Pugliatti** fisso martedì su W",
                            icon="🔒",
                        )

                with _tab_col:
                    _sel_doc = st.selectbox("Seleziona medico", _all_docs_list, key="pool_col_doc_sel")
                    if _sel_doc and _sel_doc in _draft_doctors:
                        _doc_cols = set(_draft_doctors[_sel_doc].get("columns") or [])
                        st.markdown(f"**Colonne assegnate a {_sel_doc}** — clicca per aggiungere/rimuovere")
                        _col_names = cfg_admin.get("columns", {})
                        _cols_per_row = 4
                        _col_items = sorted(_col_names.items())
                        for _ci in range(0, len(_col_items), _cols_per_row):
                            _chunk = _col_items[_ci:_ci + _cols_per_row]
                            _gcols = st.columns(len(_chunk))
                            for _gci, ((_col_letter, _col_name), _gc) in enumerate(zip(_chunk, _gcols)):
                                _is_on = _col_letter in _doc_cols
                                _locked = _col_letter == "C"
                                if _locked:
                                    _gc.markdown(f"{'🔵' if _is_on else '⚫'} **{_col_letter}** — {_col_name}  \n*(C gestita da Reperibilità)*")
                                else:
                                    if _gc.checkbox(f"{_col_letter} · {_col_name}", value=_is_on,
                                                    key=f"pool_col_{_sel_doc}_{_col_letter}"):
                                        _doc_cols.add(_col_letter)
                                    else:
                                        _doc_cols.discard(_col_letter)
                        _draft_doctors[_sel_doc]["columns"] = sorted(_doc_cols)

                with _tab_lim:
                    st.markdown("**Impostazioni globali per colonna**")
                    _col_settings: dict = _draft.setdefault("column_settings", {})
                    _cs_rows = []
                    for _cl in _all_cols:
                        _cs = _col_settings.get(_cl) or {}
                        _cs_rows.append({
                            "Colonna": _cl,
                            "Nome": (cfg_admin.get("columns") or {}).get(_cl, ""),
                            "Target mensile": _cs.get("monthly_target"),
                            "Spacing min (gg)": int(_cs.get("spacing_min_days", 0) or 0),
                            "Spacing pref (gg)": int(_cs.get("spacing_preferred_days", 0) or 0),
                            "Conta come": int(_cs.get("counts_as", 1) if _cl != "C" else 0),
                        })
                    _cs_df = _pd_pool.DataFrame(_cs_rows)
                    _edited_cs = st.data_editor(
                        _cs_df,
                        column_config={
                            "Colonna": st.column_config.TextColumn("Col", disabled=True),
                            "Nome": st.column_config.TextColumn("Nome", disabled=True),
                            "Target mensile": st.column_config.NumberColumn("Target", min_value=0, max_value=31, step=1),
                            "Spacing min (gg)": st.column_config.NumberColumn("Spacing min", min_value=0, max_value=30, step=1),
                            "Spacing pref (gg)": st.column_config.NumberColumn("Spacing pref", min_value=0, max_value=30, step=1, help="Solo per J — soft preference"),
                            "Conta come": st.column_config.NumberColumn("Conta come", min_value=0, max_value=4, step=1, help="C=0 (bloccato), J=2, altri=1"),
                        },
                        hide_index=True, use_container_width=True, key="pool_cs_editor", num_rows="fixed",
                    )
                    for _, _row in _edited_cs.iterrows():
                        _cl = _row["Colonna"]
                        _csd = _col_settings.setdefault(_cl, {})
                        _mt = _row["Target mensile"]
                        _csd["monthly_target"] = int(_mt) if _mt is not None and not _pd_pool.isna(_mt) else None
                        _csd["spacing_min_days"] = int(_row["Spacing min (gg)"] or 0)
                        _csd["spacing_preferred_days"] = int(_row["Spacing pref (gg)"] or 0)
                        _csd["counts_as"] = 0 if _cl == "C" else int(_row["Conta come"] or 1)

                    st.divider()
                    st.markdown("**Override quota per singolo medico**")
                    _ov_rows = []
                    for _dname, _dc in _draft_doctors.items():
                        for _col, _ov in (_dc.get("column_overrides") or {}).items():
                            if not isinstance(_ov, dict):
                                continue
                            _ov_rows.append({
                                "Medico": _dname, "Colonna": _col,
                                "Quota mensile": _ov.get("monthly_quota"),
                                "Tipo": _ov.get("quota_type", "fixed"),
                                "Notti weekend": _ov.get("weekend_nights", True) if _col == "J" else None,
                            })
                    _ov_df = _pd_pool.DataFrame(_ov_rows) if _ov_rows else _pd_pool.DataFrame(
                        columns=["Medico", "Colonna", "Quota mensile", "Tipo", "Notti weekend"]
                    )
                    _edited_ov = st.data_editor(
                        _ov_df,
                        column_config={
                            "Medico": st.column_config.SelectboxColumn("Medico", options=_all_docs_list),
                            "Colonna": st.column_config.SelectboxColumn("Colonna", options=_all_cols),
                            "Quota mensile": st.column_config.NumberColumn("Quota", min_value=0, max_value=31),
                            "Tipo": st.column_config.SelectboxColumn("Tipo", options=["fixed", "max", "min"]),
                            "Notti weekend": st.column_config.CheckboxColumn("Notti weekend (J)", help="Solo per col. J"),
                        },
                        hide_index=True, use_container_width=True, key="pool_ov_editor", num_rows="dynamic",
                    )
                    _new_overrides: dict[str, dict] = {}
                    for _, _row in _edited_ov.iterrows():
                        _dname = _row.get("Medico")
                        _col = _row.get("Colonna")
                        if not _dname or not _col or _dname not in _draft_doctors:
                            continue
                        _ov_entry: dict = {}
                        _mq = _row.get("Quota mensile")
                        if _mq is not None and not _pd_pool.isna(_mq):
                            _ov_entry["monthly_quota"] = int(_mq)
                            _ov_entry["quota_type"] = str(_row.get("Tipo") or "fixed")
                        if _col == "J":
                            _wn = _row.get("Notti weekend")
                            if _wn is not None and not _pd_pool.isna(_wn):
                                _ov_entry["weekend_nights"] = bool(_wn)
                        if _ov_entry:
                            _new_overrides.setdefault(_dname, {})[_col] = _ov_entry
                    for _dname in _draft_doctors:
                        _draft_doctors[_dname]["column_overrides"] = _new_overrides.get(_dname, {})

                with _tab_serv:
                    st.markdown("**Combinazioni same-day**")
                    _combos: list = _draft.setdefault("service_combinations", [])
                    _combo_rows = [
                        {"Col 1": c["columns"][0], "Col 2": c["columns"][1], "Modalità": c["mode"]}
                        for c in _combos if len(c.get("columns", [])) == 2
                    ]
                    _combo_df = _pd_pool.DataFrame(_combo_rows) if _combo_rows else _pd_pool.DataFrame(columns=["Col 1", "Col 2", "Modalità"])
                    _edited_combo = st.data_editor(
                        _combo_df,
                        column_config={
                            "Col 1": st.column_config.SelectboxColumn("Col 1", options=_all_cols),
                            "Col 2": st.column_config.SelectboxColumn("Col 2", options=_all_cols),
                            "Modalità": st.column_config.SelectboxColumn("Modalità", options=["always", "fallback", "preferred"]),
                        },
                        hide_index=True, use_container_width=True, key="pool_combo_editor", num_rows="dynamic",
                    )
                    _new_combos = []
                    for _, _row in _edited_combo.iterrows():
                        _c1, _c2, _mode = _row.get("Col 1"), _row.get("Col 2"), _row.get("Modalità")
                        if _c1 and _c2 and _mode:
                            _new_combos.append({"columns": [str(_c1), str(_c2)], "same_day": True, "mode": str(_mode)})
                    _draft["service_combinations"] = _new_combos

                    st.divider()
                    st.markdown("**Servizi critici** — fallback se pool primario esaurito")
                    _critical: dict = _draft.setdefault("critical_services", {})
                    _crit_rows = []
                    for _col, _spec in _critical.items():
                        _fb = _spec.get("fallback", "any")
                        _crit_rows.append({"Colonna": _col, "Fallback": "any" if _fb == "any" else ", ".join(_fb)})
                    _crit_df = _pd_pool.DataFrame(_crit_rows) if _crit_rows else _pd_pool.DataFrame(columns=["Colonna", "Fallback"])
                    _edited_crit = st.data_editor(
                        _crit_df,
                        column_config={
                            "Colonna": st.column_config.SelectboxColumn("Colonna", options=_all_cols),
                            "Fallback": st.column_config.TextColumn("Fallback", help="'any' oppure nomi separati da virgola"),
                        },
                        hide_index=True, use_container_width=True, key="pool_crit_editor", num_rows="dynamic",
                    )
                    _new_crit: dict = {}
                    for _, _row in _edited_crit.iterrows():
                        _col = _row.get("Colonna")
                        _fb_raw = str(_row.get("Fallback") or "any").strip()
                        if not _col:
                            continue
                        if _fb_raw.lower() == "any":
                            _new_crit[str(_col)] = {"fallback": "any"}
                        else:
                            _fb_list = [x.strip() for x in _fb_raw.split(",") if x.strip()]
                            if _fb_list:
                                _new_crit[str(_col)] = {"fallback": _fb_list}
                    _draft["critical_services"] = _new_crit

                st.divider()
                _col_save, _col_reset = st.columns([3, 1])
                with _col_reset:
                    if st.button("↩️ Reset draft", key="btn_pool_reset"):
                        del st.session_state["pool_cfg_draft"]
                        st.rerun()
                with _col_save:
                    if st.button("💾 Salva configurazione pool", type="primary", key="btn_pool_save"):
                        import copy as _copy_save
                        from datetime import datetime as _dt_save, timezone as _tz_save
                        _to_save = _copy_save.deepcopy(_draft)
                        _to_save["updated_at"] = _dt_save.now(_tz_save.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                        _to_save["updated_by"] = "admin"
                        _errs = _pcs_ui.validate_pool_config(_to_save)
                        if _errs:
                            st.error("Configurazione non valida:\n\n" + "\n".join(f"- {e}" for e in _errs))
                        else:
                            _ok_save, _msg_save = save_pool_config_with_retry(_to_save, _pool_cfg_sha)
                            if _ok_save:
                                st.success(_msg_save, icon="✅")
                                del st.session_state["pool_cfg_draft"]
                                st.rerun()
                            else:
                                st.error(_msg_save)

        # ── Memoria storica turni ────────────────────────────────────────
        with st.expander("📊 Memoria storica turni", expanded=False):
            st.info(
                "Carica i file Excel **definitivi** dei mesi precedenti per costruire "
                "una memoria storica. Il solver userà questi dati per bilanciare le "
                "quote tra i mesi.",
                icon="🧠",
            )
            _hist_data, _hist_sha = _load_shift_history()

            with st.expander("📤 Carica mese definitivo", expanded=False):
                _hist_upload = st.file_uploader(
                    "File Excel turni definitivo", type=["xlsx"], key="hist_upload",
                    help="Il file Excel finale (dopo le modifiche del primario) di un mese passato.",
                )
                if _hist_upload is not None:
                    if st.button("📥 Importa nel storico", key="btn_import_hist"):
                        _tmp = Path(tempfile.gettempdir()) / f"hist_{int(time.time())}.xlsx"
                        try:
                            _tmp.write_bytes(_hist_upload.getvalue())
                            _parsed = sh.parse_finalized_xlsx(str(_tmp))
                            _ml = _parsed["month_label"]
                            if not _ml:
                                st.error("Impossibile determinare il mese dal file Excel.")
                            else:
                                _valid_docs = set(doctors) if doctors else None
                                _ms = sh.compute_doctor_stats(_parsed, valid_doctors=_valid_docs)
                                _last_night = []
                                if _parsed["days"]:
                                    _last_night = _parsed["days"][-1].get("assignments", {}).get("J", [])
                                _ms["_meta"] = {"last_day_night_doctors": _last_night}
                                _hist_data[_ml] = _ms
                                if _save_shift_history(_hist_data, _hist_sha):
                                    st.success(f"✅ Mese **{_ml}** importato ({len(_parsed['days'])} giorni)")
                                    st.rerun()
                        except Exception as _e:
                            st.error(f"Errore parsing: {_e}")
                        finally:
                            if _tmp.exists():
                                _tmp.unlink()

            if _hist_data:
                _sorted_months = sorted(_hist_data.keys())
                st.caption(f"Mesi in memoria: {', '.join(_sorted_months)}")
                _agg_hist = sh.aggregate_multi_month(_hist_data)

                with st.expander("📋 Tabella riepilogativa", expanded=True):
                    _tab_cum, _tab_mese = st.tabs(["Cumulativo", "Per mese"])
                    with _tab_cum:
                        _rows_hist = []
                        for _doc in sorted(_agg_hist.keys()):
                            _ds = _agg_hist[_doc]
                            _j = _ds.get("J", {}); _c = _ds.get("C", {})
                            _h = _ds.get("H", {}); _i = _ds.get("I", {})
                            _rows_hist.append({
                                "Medico": _doc, "Mesi": _ds.get("_months_counted", 0),
                                "Notti (J)": _j.get("total", 0) if isinstance(_j, dict) else 0,
                                "Notti Sab": _j.get("sabati", 0) if isinstance(_j, dict) else 0,
                                "Notti Dom": _j.get("domeniche", 0) if isinstance(_j, dict) else 0,
                                "Reperibilità (C)": _c.get("total", 0) if isinstance(_c, dict) else 0,
                                "Festivi (D/E/H/I)": _ds.get("_festivi_DE_HI", 0),
                                "Domeniche": _ds.get("_domeniche", 0), "Sabati": _ds.get("_sabati", 0),
                                "H pom. (fer.)": _h.get("feriali", 0) if isinstance(_h, dict) else 0,
                                "I pom. (fer.)": _i.get("feriali", 0) if isinstance(_i, dict) else 0,
                            })
                        _df_hist = pd.DataFrame(_rows_hist)
                        st.dataframe(_df_hist, use_container_width=True, hide_index=True)
                    with _tab_mese:
                        _sel_mese = st.selectbox("Mese", _sorted_months, key="hist_tab_mese", index=len(_sorted_months)-1)
                        _ms_sel = _hist_data[_sel_mese]
                        _rows_mese = []
                        for _doc in sorted(k for k in _ms_sel.keys() if k != "_meta"):
                            _ds = _ms_sel[_doc]
                            _j = _ds.get("J", {}); _c = _ds.get("C", {})
                            _h = _ds.get("H", {}); _i = _ds.get("I", {})
                            _rows_mese.append({
                                "Medico": _doc,
                                "Notti (J)": _j.get("total", 0) if isinstance(_j, dict) else 0,
                                "Notti Sab": _j.get("sabati", 0) if isinstance(_j, dict) else 0,
                                "Notti Dom": _j.get("domeniche", 0) if isinstance(_j, dict) else 0,
                                "Reperibilità (C)": _c.get("total", 0) if isinstance(_c, dict) else 0,
                                "Festivi (D/E/H/I)": _ds.get("_festivi_DE_HI", 0),
                                "Domeniche": _ds.get("_domeniche", 0), "Sabati": _ds.get("_sabati", 0),
                                "H pom. (fer.)": _h.get("feriali", 0) if isinstance(_h, dict) else 0,
                                "I pom. (fer.)": _i.get("feriali", 0) if isinstance(_i, dict) else 0,
                            })
                        st.dataframe(pd.DataFrame(_rows_mese), use_container_width=True, hide_index=True)

                with st.expander("📈 Grafici", expanded=False):
                    if _rows_hist:
                        _graf_scelta = st.selectbox("Grafico", [
                            "Notti totali (cumulativo)", "Notti: feriali / sabato / domenica (cumulativo)",
                            "Domeniche lavorate (cumulativo)", "Reperibilità (cumulativo)",
                            "Evoluzione notti mese per mese",
                        ], key="hist_graf_sel")
                        if _graf_scelta == "Notti totali (cumulativo)":
                            _fig = px.bar(_df_hist, x="Medico", y="Notti (J)", title="Notti totali per medico (cumulativo)", color="Notti (J)", color_continuous_scale="Reds")
                            _fig.update_layout(xaxis_tickangle=-45, height=400); st.plotly_chart(_fig, use_container_width=True)
                        elif _graf_scelta == "Notti: feriali / sabato / domenica (cumulativo)":
                            _df_j2 = pd.DataFrame([{"Medico": r["Medico"], "Feriali": r["Notti (J)"]-r["Notti Sab"]-r["Notti Dom"], "Sabato": r["Notti Sab"], "Domenica": r["Notti Dom"]} for r in _rows_hist])
                            _fig = px.bar(_df_j2, x="Medico", y=["Feriali","Sabato","Domenica"], title="Notti: distribuzione feriali/sabato/domenica", barmode="stack")
                            _fig.update_layout(xaxis_tickangle=-45, height=400); st.plotly_chart(_fig, use_container_width=True)
                        elif _graf_scelta == "Domeniche lavorate (cumulativo)":
                            _fig = px.bar(_df_hist, x="Medico", y="Domeniche", title="Domeniche lavorate per medico (cumulativo)", color="Domeniche", color_continuous_scale="Blues")
                            _fig.update_layout(xaxis_tickangle=-45, height=400); st.plotly_chart(_fig, use_container_width=True)
                        elif _graf_scelta == "Reperibilità (cumulativo)":
                            _fig = px.bar(_df_hist, x="Medico", y="Reperibilità (C)", title="Reperibilità per medico (cumulativo)", color="Reperibilità (C)", color_continuous_scale="Greens")
                            _fig.update_layout(xaxis_tickangle=-45, height=400); st.plotly_chart(_fig, use_container_width=True)
                        elif _graf_scelta == "Evoluzione notti mese per mese":
                            if len(_sorted_months) > 1:
                                _evo = []
                                for _ml2 in _sorted_months:
                                    for _doc2, _ds2 in _hist_data[_ml2].items():
                                        if _doc2 == "_meta": continue
                                        _j2 = _ds2.get("J", {})
                                        _evo.append({"Mese": _ml2, "Medico": _doc2, "Notti": _j2.get("total", 0) if isinstance(_j2, dict) else 0})
                                _fig = px.line(pd.DataFrame(_evo), x="Mese", y="Notti", color="Medico", title="Evoluzione notti per medico", markers=True)
                                _fig.update_layout(height=400); st.plotly_chart(_fig, use_container_width=True)
                            else:
                                st.info("Servono almeno 2 mesi per il grafico di evoluzione.")

                with st.expander("🗑️ Rimuovi mese dallo storico", expanded=False):
                    _month_del = st.selectbox("Seleziona mese da rimuovere", _sorted_months, key="hist_del")
                    if st.button("Rimuovi", key="btn_del_hist"):
                        if _month_del in _hist_data:
                            del _hist_data[_month_del]
                            if _save_shift_history(_hist_data, _hist_sha):
                                st.success(f"Mese {_month_del} rimosso."); st.rerun()
            else:
                st.caption("Nessun mese caricato nella memoria storica.")

        # ── Contatore turni universitari ──────────────────────────────────
        with st.expander("🎓 Contatore turni universitari", expanded=False):
            try:
                gc_u = (cfg_admin.get("global_constraints") or {})
                uni_docs = gc_u.get("university_doctors") or {}
                uni_ratio = float(gc_u.get("university_ratio", 0.6))
                if uni_docs:
                    import calendar as _cal
                    _cont_col1, _cont_col2 = st.columns(2)
                    with _cont_col1:
                        _cont_year = st.number_input("Anno", min_value=2025, max_value=2035, value=date.today().year, step=1, key="cont_year")
                    with _cont_col2:
                        _cont_month = st.number_input("Mese", min_value=1, max_value=12, value=date.today().month, step=1, key="cont_month")
                    _, n_days = _cal.monthrange(int(_cont_year), int(_cont_month))
                    _holidays = tg.italy_public_holidays(int(_cont_year))
                    working = sum(1 for d in range(1, n_days+1)
                        if (dt_date := date(int(_cont_year), int(_cont_month), d)).weekday() < 6
                        and dt_date not in _holidays)
                    _cont_mk = f"{int(_cont_year)}-{int(_cont_month):02d}"
                    st.markdown(f"**{_cont_mk}** — giorni lavorativi lun-sab (esclusi festivi): **{working}**")
                    st.markdown(f"Rapporto universitari: **{int(uni_ratio*100)}%** → target = round({working} × {uni_ratio}) = **{round(working * uni_ratio)}**")
                    rows_u = []
                    for doc_raw, dcfg in uni_docs.items():
                        night_double = bool((dcfg or {}).get("night_counts_double", False))
                        target = round(working * uni_ratio)
                        rows_u.append({"Medico": doc_raw, "Notte vale doppio": "✓ (2 turni)" if night_double else "✗ (1 turno)", "Target turni pesati": target, "Max consentito": target+1, "Note": "Indisponibilità NON riducono il target"})
                    st.dataframe(pd.DataFrame(rows_u), use_container_width=True, hide_index=True)
                else:
                    st.info("Nessun medico universitario configurato nel YAML.")
            except Exception as e:
                st.warning(f"Errore contatore universitari: {e}")

        # ── Migrazioni (condizionali) ─────────────────────────────────────
        _mg = _github_cfg()
        _unavail_dir_files = github_utils.list_dir(
            owner=_mg["owner"], repo=_mg["repo"],
            path=_unavail_per_doctor_dir(),
            token=_mg["token"], branch=_mg.get("branch", "main"),
        )
        _unavail_already_migrated = len(_unavail_dir_files) > 0

        if _unavail_already_migrated:
            with st.expander("✅ Indisponibilità — file per-medico già presenti", expanded=False):
                st.success(f"Trovati {len(_unavail_dir_files)} file in `{_unavail_per_doctor_dir()}/`. Migrazione già effettuata.", icon="✅")
                st.caption("Se necessario puoi ripetere la migrazione qui sotto.")
                _do_migrate = st.button("Ripeti migrazione indisponibilità", key="btn_migrate_unavail")
        else:
            st.markdown("#### Migrazione indisponibilità al nuovo formato")
            st.warning("**Azione richiesta (una-tantum).** Il sistema ora salva un file CSV separato per ogni medico, eliminando i conflitti di salvataggio concorrente. Clicca il bottone qui sotto per copiare i dati storici dal vecchio CSV ai file per-medico.", icon="⚠️")
            _mg_col1, _mg_col2 = st.columns([1, 3])
            with _mg_col1:
                _do_migrate = st.button("Esegui migrazione", key="btn_migrate_unavail", type="primary", use_container_width=True)
            with _mg_col2:
                st.caption(f"Legge `unavailability_store.csv` → scrive file per-medico in `{_unavail_per_doctor_dir()}/`. Il vecchio file resta intatto come backup.")
        if _do_migrate:
            if not _mg.get("path"):
                st.error("Chiave `path` non trovata nei secrets — migrazione non possibile.")
            else:
                try:
                    _leg_gf = github_utils.get_file(owner=_mg["owner"], repo=_mg["repo"], path=_mg["path"], token=_mg["token"], branch=_mg.get("branch", "main"))
                    if _leg_gf is None or not (_leg_gf.text or "").strip():
                        st.warning("CSV aggregato vuoto o non trovato — nessun dato da migrare.")
                    else:
                        from collections import defaultdict as _dd
                        _all = ustore.load_store(_leg_gf.text)
                        _by_doc: dict[str, list[dict]] = _dd(list)
                        for _r in _all:
                            _by_doc[_r["doctor"]].append(_r)
                        _prog = st.progress(0)
                        _docs_list = list(_by_doc.keys()); _ok = 0
                        for _i, _doc in enumerate(_docs_list):
                            _doc_path = _doctor_unavail_path(_doc)
                            _ex_gf = github_utils.get_file(owner=_mg["owner"], repo=_mg["repo"], path=_doc_path, token=_mg["token"], branch=_mg.get("branch", "main"))
                            github_utils.put_file(owner=_mg["owner"], repo=_mg["repo"], path=_doc_path, token=_mg["token"], branch=_mg.get("branch", "main"), sha=_ex_gf.sha if _ex_gf else None, message=f"migrate: unavailability for {_doc}", text=ustore.to_csv(_by_doc[_doc]))
                            _ok += 1; _prog.progress((_i+1)/len(_docs_list))
                        st.success(f"Migrazione completata: {_ok} medici migrati in `{_unavail_per_doctor_dir()}/`.")
                except Exception as _me:
                    st.error(f"Errore migrazione: {_me}")

        _mga = _github_cfg()
        _avail_dir_files = github_utils.list_dir(owner=_mga["owner"], repo=_mga["repo"], path=_avail_per_doctor_dir(), token=_mga["token"], branch=_mga.get("branch", "main"))
        _avail_already_migrated = len(_avail_dir_files) > 0

        if _avail_already_migrated:
            with st.expander("✅ Disponibilità — file per-medico già presenti", expanded=False):
                st.success(f"Trovati {len(_avail_dir_files)} file in `{_avail_per_doctor_dir()}/`. Migrazione già effettuata.", icon="✅")
                st.caption("Se necessario puoi ripetere la migrazione qui sotto.")
                _do_migrate_avail = st.button("Ripeti migrazione disponibilità", key="btn_migrate_avail")
        else:
            st.markdown("#### Migrazione preferenze disponibilità al nuovo formato")
            st.warning(f"**Azione richiesta (una-tantum).** Divide `availability_store.csv` in file per-medico nella directory `{_avail_per_doctor_dir()}/`, eliminando i conflitti concorrenti.", icon="⚠️")
            _mg2_col1, _mg2_col2 = st.columns([1, 3])
            with _mg2_col1:
                _do_migrate_avail = st.button("Esegui migrazione disponibilità", key="btn_migrate_avail", type="primary", use_container_width=True)
            with _mg2_col2:
                st.caption(f"Legge `availability_store.csv` → scrive file per-medico in `{_avail_per_doctor_dir()}/`. Il vecchio file resta intatto come backup.")
        if _do_migrate_avail:
            try:
                _avail_legacy_path = _mga.get("availability_path", "data/availability_store.csv")
                _leg_avail_gf = github_utils.get_file(owner=_mga["owner"], repo=_mga["repo"], path=_avail_legacy_path, token=_mga["token"], branch=_mga.get("branch", "main"))
                if _leg_avail_gf is None or not (_leg_avail_gf.text or "").strip():
                    st.warning("CSV aggregato disponibilità vuoto o non trovato — nessun dato da migrare.")
                else:
                    from collections import defaultdict as _dd2
                    _all_avail = ustore.load_store(_leg_avail_gf.text)
                    _by_doc_avail: dict[str, list[dict]] = _dd2(list)
                    for _r in _all_avail:
                        _by_doc_avail[_r["doctor"]].append(_r)
                    _prog2 = st.progress(0); _docs2 = list(_by_doc_avail.keys()); _ok2 = 0
                    for _i2, _doc2 in enumerate(_docs2):
                        _dp = _doctor_avail_path(_doc2)
                        _ex2 = github_utils.get_file(owner=_mga["owner"], repo=_mga["repo"], path=_dp, token=_mga["token"], branch=_mga.get("branch", "main"))
                        github_utils.put_file(owner=_mga["owner"], repo=_mga["repo"], path=_dp, token=_mga["token"], branch=_mga.get("branch", "main"), sha=_ex2.sha if _ex2 else None, message=f"migrate: availability for {_doc2}", text=ustore.to_csv(_by_doc_avail[_doc2]))
                        _ok2 += 1; _prog2.progress((_i2+1)/len(_docs2))
                    st.success(f"Migrazione disponibilità completata: {_ok2} medici → `{_avail_per_doctor_dir()}/`.")
            except Exception as _me2:
                st.error(f"Errore migrazione disponibilità: {_me2}")

        st.stop()  # Fine ramo Configurazione

    # ── Ramo Genera turni ────────────────────────────────────────────────
    # Step 1: Periodo
    st.markdown("### 1) Periodo")
    today = date.today()
    colA, colB, colC = st.columns([1, 1, 2])
    with colA:
        year = st.number_input("Anno", min_value=2025, max_value=2035, value=today.year, step=1)
    with colB:
        month = st.number_input("Mese", min_value=1, max_value=12, value=today.month, step=1)
    mk = f"{int(year)}-{int(month):02d}"
    st.caption(f"Stai generando: **{mk}**")

    # Step 2: Indisponibilità
    st.markdown("### 2) Indisponibilità")
    if "admin_unav_mode" not in st.session_state:
        st.session_state["admin_unav_mode"] = "Usa archivio (privacy)"
    unav_mode = st.radio(
        "Fonte indisponibilità",
        ["Nessuna", "Carica file manuale", "Usa archivio (privacy)"],
        key="admin_unav_mode",
        horizontal=True,
        help="Puoi caricare un file manuale, oppure usare l’archivio compilato dai medici.",
    )
    unav_upload = None
    if unav_mode == "Carica file manuale":
        unav_upload = st.file_uploader("Carica indisponibilità (xlsx/csv/tsv)", type=["xlsx", "csv", "tsv"])
    use_archive = (unav_mode == "Usa archivio (privacy)")

    with st.expander("📜 Log inserimenti/modifiche indisponibilità (Audit)", expanded=False):
        st.write(
            "Questo log registra chi ha inserito/modificato indisponibilità, con timestamp e conteggi. "
            "È utile per tracciare le modifiche mese per mese."
        )

        cL1, cL2, cL3 = st.columns([1, 1, 2])
        with cL1:
            audit_year = st.number_input(
                "Anno log", min_value=2025, max_value=2035, value=int(year), step=1, key="audit_year",
            )
        with cL2:
            audit_month = st.number_input(
                "Mese log", min_value=1, max_value=12, value=int(month), step=1, key="audit_month",
            )
        with cL3:
            mk_log = f"{int(audit_year)}-{int(audit_month):02d}"
            st.caption(f"Log selezionato: **{mk_log}**")

        try:
            audit_text = load_audit_log_text_from_github(mk_log)
        except Exception as e:
            audit_text = None
            st.error(f"Errore lettura audit log da GitHub: {e}")

        if not audit_text or not str(audit_text).strip():
            st.info("Nessun audit log trovato per questo mese.")
        else:
            st.download_button(
                "⬇️ Scarica audit log (CSV)", data=str(audit_text).encode("utf-8"),
                file_name=f"unavailability_audit_{mk_log}.csv", mime="text/csv",
                key=f"dl_audit_csv_{mk_log}",
            )
            try:
                df_audit = pd.read_csv(io.StringIO(audit_text))
            except Exception:
                df_audit = None
            if df_audit is not None and not df_audit.empty:
                try:
                    if "ts_utc" in df_audit.columns:
                        df_audit = df_audit.sort_values("ts_utc", ascending=False)
                except Exception:
                    pass
                doctor_filter = "Tutti"
                if "doctor" in df_audit.columns:
                    doctors_in_log = sorted([str(x) for x in df_audit["doctor"].dropna().unique().tolist() if str(x).strip()])
                    doctor_filter = st.selectbox("Filtro medico", ["Tutti"] + doctors_in_log, index=0, key=f"audit_filter_{mk_log}")
                df_preview = df_audit if doctor_filter == "Tutti" else df_audit[df_audit["doctor"] == doctor_filter]
                try:
                    xlsx_bytes = audit_df_to_excel_bytes(df_audit, sheet_name=f"audit_{mk_log}")
                    st.download_button("⬇️ Scarica audit log (Excel)", data=xlsx_bytes, file_name=f"unavailability_audit_{mk_log}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_audit_xlsx_{mk_log}")
                except Exception:
                    pass
                st.markdown("**Anteprima**")
                st.dataframe(df_preview.head(200), use_container_width=True, hide_index=True)
                st.caption("Mostro al massimo 200 righe. Per analisi completa usa il download.")

    # Step 3: Vincolo post-notte (carryover)
    st.markdown("### 3) Vincolo post-notte a cavallo mese")
    st.info("Inserire medico che ha fatto turno di NOTTE l’ultimo giorno del mese precedente.", icon="💡")

    # Auto-carryover da storico: chi ha fatto notte l’ultimo giorno del mese precedente
    _carry_default = []
    try:
        from datetime import datetime as _dt, timedelta as _td
        _mk_prev = (_dt(int(year), int(month), 1) - _td(days=1))
        _mk_minus1 = f"{_mk_prev.year:04d}-{_mk_prev.month:02d}"
        _hd_carry, _ = _load_shift_history()
        if _hd_carry:
            _last_month = sorted(_hd_carry.keys())[-1]
            if _last_month == _mk_minus1:
                _meta = _hd_carry[_last_month].get("_meta", {})
                _carry_default = [
                    d for d in _meta.get("last_day_night_doctors", [])
                    if d in doctors
                ]
    except Exception:
        pass

    manual_block = st.multiselect(
        "Seleziona medico/i da bloccare il Giorno 1",
        doctors,
        default=_carry_default,
        help="Chi ha fatto NOTTE l’ultimo giorno del mese precedente. "
             "Pre-compilato dallo storico se disponibile (solo se il mese "
             "precedente è nello storico).",
    )

    carryover_by_month = {}

    if manual_block:
        carryover_by_month.setdefault(mk, {})
        carryover_by_month[mk].setdefault("blocked_day1_doctors", [])
        for d in manual_block:
            if d not in carryover_by_month[mk]["blocked_day1_doctors"]:
                carryover_by_month[mk]["blocked_day1_doctors"].append(d)

    st.divider()

    # ── Step 4: Assegnazioni fisse ────────────────────────────────────────────
    st.markdown("### 4) Assegnazioni fisse (opzionale)")
    st.info(
        "Fissa un medico in una colonna specifica in un giorno specifico. "
        "Il solver rispetterà questi vincoli in modo OBBLIGATORIO.",
        icon="📌",
    )

    SHIFT_LABELS_ADMIN = {
        "Notte (J)": "J",
        "UTIC mattina (D)": "D",
        "Supporto 118 (F)": "F",
        "Cardiologia mattina (E)": "E",
        "Riabilitazione (G)": "G",
        "UTIC pomeriggio (H)": "H",
        "Cardiologia pomeriggio (I)": "I",
        "Letto (K)": "K",
        "Padiglioni (L)": "L",
        "ECO base (Q)": "Q",
        "ECOSTRESS/R (R)": "R",
        "Ecosala (S)": "S",
        "Interni (T)": "T",
        "Contr.PM (U)": "U",
        "Sala PM (V)": "V",
        "Ergometria (W)": "W",
        "Vascolare (Z)": "Z",
        "Holter/AB (AB)": "AB",
        "Scintigrafia (AC)": "AC",
        "Reperibilità (C)": "C",
    }

    fa_key = f"fixed_assign_{mk}"
    if fa_key not in st.session_state:
        st.session_state[fa_key] = []

    fa_c1, fa_c2, _ = st.columns([1, 1, 4])
    with fa_c1:
        if st.button("➕ Aggiungi assegnazione fissa", key="fa_add"):
            import calendar
            first_of_month = date(int(year), int(month), 1)
            st.session_state[fa_key].append({
                "id": str(uuid.uuid4()),
                "doctor": doctors[0] if doctors else "",
                "date": first_of_month,
                "column_label": list(SHIFT_LABELS_ADMIN.keys())[0],
            })
            st.rerun()
    with fa_c2:
        if st.button("🗑 Rimuovi tutte", key="fa_clear") and st.session_state[fa_key]:
            st.session_state[fa_key] = []
            st.rerun()

    updated_fa = []
    for fa_r in st.session_state.get(fa_key, []):
        import calendar
        fa_cols = st.columns([2, 2, 2, 0.5])
        with fa_cols[0]:
            fa_doc = st.selectbox("Medico", doctors,
                                  index=doctors.index(fa_r["doctor"]) if fa_r["doctor"] in doctors else 0,
                                  key=f"fa_doc_{fa_r['id']}")
        with fa_cols[1]:
            last_day_month = calendar.monthrange(int(year), int(month))[1]
            fa_date = st.date_input("Giorno",
                                    value=fa_r.get("date") or date(int(year), int(month), 1),
                                    min_value=date(int(year), int(month), 1),
                                    max_value=date(int(year), int(month), last_day_month),
                                    key=f"fa_date_{fa_r['id']}",
                                    format="DD/MM/YYYY")
        with fa_cols[2]:
            fa_col_label = st.selectbox("Turno/Colonna",
                                        list(SHIFT_LABELS_ADMIN.keys()),
                                        index=list(SHIFT_LABELS_ADMIN.keys()).index(fa_r.get("column_label", list(SHIFT_LABELS_ADMIN.keys())[0])),
                                        key=f"fa_col_{fa_r['id']}")
        with fa_cols[3]:
            fa_del = st.button("🗑", key=f"fa_del_{fa_r['id']}")
        if not fa_del:
            updated_fa.append({"id": fa_r["id"], "doctor": fa_doc, "date": fa_date, "column_label": fa_col_label})
    st.session_state[fa_key] = updated_fa

    # Converti in formato per il solver
    fixed_assignments_list = [
        {"doctor": r["doctor"], "date": str(r["date"]), "column": SHIFT_LABELS_ADMIN[r["column_label"]]}
        for r in st.session_state.get(fa_key, [])
        if r.get("doctor") and r.get("date") and r.get("column_label")
    ]
    if fixed_assignments_list:
        st.caption(f"Assegnazioni fisse attive: {len(fixed_assignments_list)} → " +
                   ", ".join([f"{r['doctor']} il {r['date']} in {r['column']}" for r in fixed_assignments_list]))

    st.divider()

    # ── Step 5: Turno doppio Sala PM (V) ──────────────────────────────────────
    st.markdown("### 5) Turno doppio Sala PM (V) — eccezioni settimanali")
    st.info(
        "Di default ogni **venerdì** ha il turno doppio in V (Crea + Dattilo o Allegra). "
        "Puoi spostare il doppio su lunedì o mercoledì, oppure scegliere **nessun doppio**: "
        "in quel caso tutti e tre i giorni (lun/mer/ven) avranno un singolo medico tra Dattilo, Crea e Allegra.",
        icon="🔬",
    )

    import calendar as _calendar
    _yy_v, _mm_v = int(year), int(month)
    _n_days_v = _calendar.monthrange(_yy_v, _mm_v)[1]
    # Raccogli tutti i giorni del mese per V (lun=0, mer=2, ven=4)
    _V_WEEKDAYS = {0: "Lunedì", 2: "Mercoledì", 4: "Venerdì"}
    # Raggruppa per settimana ISO
    _weeks_v: dict = {}  # iso_week -> {dow_int: date}
    for _d in range(1, _n_days_v + 1):
        _dd = date(_yy_v, _mm_v, _d)
        _wd = _dd.weekday()
        if _wd in _V_WEEKDAYS:
            _iso_w = _dd.isocalendar()[:2]
            _weeks_v.setdefault(_iso_w, {})[_wd] = _dd

    _v_double_key = f"v_double_{mk}"
    if _v_double_key not in st.session_state:
        st.session_state[_v_double_key] = {}  # iso_week -> selected date (or None = venerdì default)

    _v_double_overrides_list = []  # date ISO strings
    _any_v_override = False
    for _iso_w in sorted(_weeks_v.keys()):
        _wdays = _weeks_v[_iso_w]
        # Venerdì di questa settimana (potrebbe non essere nel mese)
        _fri = _wdays.get(4)
        # Giorni alternativi (lun, mer) presenti nel mese
        _alt_days = {_wd: _dt for _wd, _dt in _wdays.items() if _wd != 4}
        if not _alt_days:
            # Solo venerdì in questo mese per questa settimana → niente da scegliere
            continue
        # Opzioni: venerdì (default) + giorni alternativi + nessun doppio
        _label_default = f"Venerdì {_fri.strftime('%d/%m') if _fri else '(fuori mese)'} — doppio (default)"
        _label_no_double = "Nessun doppio questa settimana — tutti singoli"
        _SENTINEL_NO_DOUBLE = "NO_DOUBLE"
        _options = {_label_default: None}
        for _wd_alt, _dt_alt in sorted(_alt_days.items()):
            _options[f"{_V_WEEKDAYS[_wd_alt]} {_dt_alt.strftime('%d/%m')} — doppio (invece del venerdì)"] = _dt_alt
        _options[_label_no_double] = _SENTINEL_NO_DOUBLE

        _prev = st.session_state[_v_double_key].get(str(_iso_w))
        _prev_label = next((lb for lb, val in _options.items() if val == _prev), _label_default)
        _sel_label = st.selectbox(
            f"Settimana {_iso_w[1]} ({min(_wdays.values()).strftime('%d/%m')}–{max(_wdays.values()).strftime('%d/%m')})",
            list(_options.keys()),
            index=list(_options.keys()).index(_prev_label),
            key=f"v_double_sel_{_iso_w[0]}_{_iso_w[1]}",
        )
        _sel_date = _options[_sel_label]
        st.session_state[_v_double_key][str(_iso_w)] = _sel_date
        if _sel_date == _SENTINEL_NO_DOUBLE:
            _v_double_overrides_list.append(f"NODOUBLE:{_iso_w[0]}:{_iso_w[1]}")
            _any_v_override = True
        elif _sel_date is not None:
            _v_double_overrides_list.append(str(_sel_date))
            _any_v_override = True

    if _any_v_override:
        _caption_parts = []
        for _ov in _v_double_overrides_list:
            if _ov.startswith("NODOUBLE:"):
                _, _yr_wk = _ov.split(":", 1)
                _yr_c, _wk_c = _yr_wk.split(":")
                _caption_parts.append(f"settimana {_wk_c}/{_yr_c} — nessun doppio")
            else:
                _caption_parts.append(f"{_ov} — doppio spostato")
        st.caption("Eccezioni attive: " + ", ".join(_caption_parts))

    st.divider()

    # ── Step 6: Giorno vuoto Notte (J) — eccezioni settimanali ───────────────
    _rJ_admin = (cfg_admin.get("rules") or {}).get("J") or {}
    if _rJ_admin.get("thursday_blank"):
        st.markdown("### 6) Giorno vuoto Notte (J) — eccezioni settimanali")
        st.info(
            "Di default ogni **giovedì** la colonna J (Notte) è vuota. "
            "Puoi spostare il giorno vuoto in un altro giorno della settimana, "
            "oppure non avere nessun giorno vuoto per settimane parziali.",
            icon="🌙",
        )

        _DOW_NAMES = {0: "Lunedì", 1: "Martedì", 2: "Mercoledì", 3: "Giovedì", 4: "Venerdì", 5: "Sabato", 6: "Domenica"}
        # Raccoglie tutte le settimane del mese con almeno un giovedì o un giorno alternativo
        _weeks_j: dict = {}  # iso_week -> {dow_int: date}
        for _d in range(1, _n_days_v + 1):
            _dd = date(_yy_v, _mm_v, _d)
            _wd = _dd.weekday()
            _iso_w = _dd.isocalendar()[:2]
            _weeks_j.setdefault(_iso_w, {})[_wd] = _dd

        _j_blank_key = f"j_blank_{mk}"
        if _j_blank_key not in st.session_state:
            st.session_state[_j_blank_key] = {}  # str(iso_week) -> date string or ""

        _j_blank_week_overrides = {}  # "YYYY-WNN" -> date ISO or None
        _any_j_override = False
        for _iso_w in sorted(_weeks_j.keys()):
            _wdays = _weeks_j[_iso_w]
            _thu = _wdays.get(3)  # giovedì
            if _thu is None and not any(d in _wdays for d in range(0, 3)):
                # Nessun giovedì e nessun giorno Mon-Wed in questo mese → salta
                continue
            # Opzioni: giovedì (default se presente), altri giorni, nessun vuoto
            _label_default = f"Giovedì {_thu.strftime('%d/%m') if _thu else '(fuori mese)'} — vuoto (default)"
            _opts: dict = {_label_default: str(_thu) if _thu else None}
            for _wd_alt in sorted([w for w in _wdays if w != 3]):
                _dt_alt = _wdays[_wd_alt]
                _opts[f"{_DOW_NAMES[_wd_alt]} {_dt_alt.strftime('%d/%m')} — vuoto"] = str(_dt_alt)
            _opts["Nessun giorno vuoto questa settimana"] = ""

            _prev_val = st.session_state[_j_blank_key].get(str(_iso_w), str(_thu) if _thu else None)
            _prev_label = next((lb for lb, val in _opts.items() if val == _prev_val), _label_default)

            _week_label = f"Settimana {_iso_w[1]} ({min(_wdays.values()).strftime('%d/%m')}–{max(_wdays.values()).strftime('%d/%m')})"
            _sel_label = st.selectbox(_week_label, list(_opts.keys()),
                                      index=list(_opts.keys()).index(_prev_label),
                                      key=f"j_blank_sel_{_iso_w[0]}_{_iso_w[1]}")
            _sel_val = _opts[_sel_label]
            st.session_state[_j_blank_key][str(_iso_w)] = _sel_val

            # Costruisci chiave "YYYY-WNN" e valore per il solver
            _wk_str = f"{_iso_w[0]}-W{_iso_w[1]:02d}"
            # Solo se diverso dal default (giovedì presente e selezionato)
            _is_default = (_thu is not None and _sel_val == str(_thu))
            if not _is_default:
                _j_blank_week_overrides[_wk_str] = _sel_val if _sel_val else None
                _any_j_override = True

        if _any_j_override:
            _desc = []
            for _wk, _bd in _j_blank_week_overrides.items():
                _desc.append(f"{_wk}: {'nessun vuoto' if _bd is None else _bd}")
            st.caption("Eccezioni J attive: " + ", ".join(_desc))
    else:
        _j_blank_week_overrides = {}

    st.divider()

    # Generate button
    generate = st.button("🚀 Genera turni", type="primary")

    if generate:
        t0 = time.time()
        status = st.status("Preparazione…", expanded=True)
        try:
            with tempfile.TemporaryDirectory() as td:
                td = Path(td)

                status.update(label="Preparazione template…", state="running")
                style_path = DEFAULT_STYLE_TEMPLATE if DEFAULT_STYLE_TEMPLATE.exists() else None
                template_path = td / f"turni_{mk}.xlsx"
                tg.create_month_template_xlsx(
                    rules_path,
                    int(year),
                    int(month),
                    out_path=template_path,
                )

                status.update(label="Carico indisponibilità…", state="running")
                unav_path = None
                if unav_mode == "Carica file manuale" and unav_upload is not None:
                    unav_path = td / "unavailability.xlsx"
                    unav_path.write_bytes(unav_upload.getvalue())
                elif use_archive:
                    # Read the archive, and re-check SHA once to minimize the chance
                    # of generating from a stale snapshot while others are saving.
                    store_rows_1, sha1 = load_store_from_github()
                    rows_month = ustore.filter_month(store_rows_1, int(year), int(month))
                    unav_path = td / "unavailability_from_store.xlsx"
                    xlsx_utils.build_unavailability_xlsx(rows_month, DEFAULT_UNAV_TEMPLATE, unav_path)

                    # Double-check solo in modalità legacy (SHA disponibile)
                    if sha1 is not None:
                        store_rows_2, sha2 = load_store_from_github()
                        if sha1 and sha2 and sha2 != sha1:
                            rows_month = ustore.filter_month(store_rows_2, int(year), int(month))
                            xlsx_utils.build_unavailability_xlsx(rows_month, DEFAULT_UNAV_TEMPLATE, unav_path)
                            st.caption("Archivio indisponibilità aggiornato durante la preparazione: ricaricata l’ultima versione.")

                    st.caption(f"Archivio indisponibilità: {len(rows_month)} righe per {mk}")

                status.update(label="Generazione turni…", state="running")
                out_path = td / f"output_{mk}.xlsx"

                # Carica preferenze di disponibilità da GitHub
                try:
                    _avail_all, _ = load_avail_store_from_github()
                    _avail_month = ustore.filter_month(_avail_all, int(year), int(month))
                    all_avail_prefs = [
                        {"doctor": r["doctor"], "date": r["date"], "shift": r["shift"],
                         "priority": r.get("priority", "media")}
                        for r in _avail_month
                    ]
                except Exception as _e:
                    all_avail_prefs = []
                    st.warning(f"Impossibile caricare preferenze da GitHub: {_e}")

                # Storico aggregato per il solver
                _hist_data2, _ = _load_shift_history()
                _hist_agg_for_solver = sh.aggregate_multi_month(_hist_data2) if _hist_data2 else None

                # Pool config overlay (se presente su GitHub)
                _pool_cfg_for_solver, _ = load_pool_config_from_github_st()

                stats, log_path = tg.generate_schedule(
                    template_xlsx=template_path,
                    rules_yml=rules_path,
                    out_xlsx=out_path,
                    unavailability_path=unav_path,
                    sheet_name=None,
                    carryover_by_month=carryover_by_month if carryover_by_month else None,
                    fixed_assignments=fixed_assignments_list if fixed_assignments_list else None,
                    availability_preferences=all_avail_prefs if all_avail_prefs else None,
                    v_double_overrides=_v_double_overrides_list if _v_double_overrides_list else None,
                    j_blank_week_overrides=_j_blank_week_overrides if _j_blank_week_overrides else None,
                    historical_stats=_hist_agg_for_solver,
                    pool_config=_pool_cfg_for_solver if _pool_cfg_for_solver else None,
                )

                status.update(label="Completato ✅", state="complete")

                # Persist outputs in session_state so that download clicks do not
                # "lose" the generated files (Streamlit re-runs the script on
                # every widget interaction).
                excel_bytes = out_path.read_bytes()
                log_bytes = None
                if log_path and Path(log_path).exists():
                    log_bytes = Path(log_path).read_bytes()

                st.session_state["last_generated"] = {
                    "mk": mk,
                    "excel_bytes": excel_bytes,
                    "log_bytes": log_bytes,
                    "stats": stats,
                    "elapsed_s": round(time.time() - t0, 2),
                    "generated_at": datetime.now().isoformat(timespec="seconds"),
                }

        except Exception:
            status.update(label="Errore ❌", state="error")
            st.error("Errore durante la generazione.")
            st.code(traceback.format_exc())

    # Downloads + summary (sticky): if a file was generated for this month, keep
    # the buttons visible even after clicking one of them.
    last = st.session_state.get("last_generated")
    if isinstance(last, dict) and last.get("mk") == mk and last.get("excel_bytes"):
        _stats = last.get("stats") if isinstance(last.get("stats"), dict) else {}
        st.success(
            f"Creato ✅ in {last.get('elapsed_s')}s | status={_stats.get('status')} | {last.get('generated_at','')}"
        )

        # If the month fell back to GREEDY, shout it loudly (otherwise users
        # may think all HARD constraints were respected).
        try:
            mstat = (_stats.get("months") or {}).get(mk, {}) or {}
            if (mstat.get("status") == "GREEDY") or (_stats.get("status") == "GREEDY"):
                err = mstat.get("solver_error") or "(motivo non disponibile)"
                st.error(
                    "⚠️ ATTENZIONE: OR-Tools non è andato a buon fine e si è attivato il fallback GREEDY. "
                    "In questa modalità alcune regole (bilanciamenti/vincoli) possono NON essere rispettate.\n\n"
                    f"Dettaglio errore: {err}"
                )
        except Exception:
            pass

        st.download_button(
            "⬇️ Scarica Excel turni",
            data=last["excel_bytes"],
            file_name=f"turni_{mk}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_xlsx_{mk}",
        )
        if last.get("log_bytes"):
            st.download_button(
                "⬇️ Scarica solver log",
                data=last["log_bytes"],
                file_name=f"solverlog_{mk}.txt",
                mime="text/plain",
                key=f"dl_log_{mk}",
            )

        # Quick, user-friendly quality panel
        st.markdown("### Controlli rapidi")
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(
                f'<div class="kpi"><b>Solver</b><br>{_stats.get("status","?")}</div>',
                unsafe_allow_html=True,
            )
        with k2:
            cdiag = _stats.get("C_reperibilita_diag") if isinstance(_stats, dict) else None
            msg = "OK" if (isinstance(cdiag, dict) and cdiag.get("status", "").startswith("OK")) else "Controllare"
            st.markdown(
                f'<div class="kpi"><b>Reperibilità (C)</b><br>{msg}</div>',
                unsafe_allow_html=True,
            )
        with k3:
            blocked = (carryover_by_month.get(mk, {}) or {}).get("blocked_day1_doctors", [])
            st.markdown(
                f'<div class="kpi"><b>Carryover</b><br>{len(blocked)} bloccati Giorno 1</div>',
                unsafe_allow_html=True,
            )

        if isinstance(_stats, dict) and _stats.get("C_reperibilita_diag"):
            with st.expander("Dettagli Reperibilità (C)"):
                st.json(_stats["C_reperibilita_diag"])
