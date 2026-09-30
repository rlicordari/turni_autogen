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
import session_leases
import unavailability_drafts as udrafts
import unavailability_service as usvc
import unavailability_receipts as receipts
import unavailability_store as ustore
import xlsx_utils
import shift_history as sh
import generation_memory as gmem

# Import generator
import turni_generator as tg

APP_BUILD = "2026-02-01-ui-v7"

# ---- Concurrency & session safety (doctor mode) ----
# - One active session per doctor: in-memory lease (session_leases), a new login
#   kicks out older sessions.
# - Saves: whole doctor-month replace with conflict check (a month changed
#   elsewhere is never overwritten), read-back verification, and a queue that
#   completes the save when GitHub is rate limited or down (unavailability_service).
# - Every GitHub call goes through github_utils' gateway (retries, pacing).

DOCTOR_SESSION_TTL_MINUTES = 20

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
        elif kind == "warning":
            st.warning(msg)
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
        "generation_memory_path": _get_secret(("GITHUB_GENERATION_MEMORY_PATH",), gmem.MEMORY_PATH_DEFAULT),
    }

# ---------------- Shift history helpers ----------------
def _generation_memory_path() -> str:
    g = _github_cfg()
    return str(
        g.get("generation_memory_path")
        or g.get("generated_memory_path")
        or g.get("generated_shift_memory_path")
        or gmem.MEMORY_PATH_DEFAULT
    )


def load_generation_memory_from_github_st() -> tuple[dict, str | None]:
    g = _github_cfg()
    return gmem.load_memory_from_github(
        owner=g["owner"],
        repo=g["repo"],
        token=g["token"],
        branch=g.get("branch", "main"),
        path=_generation_memory_path(),
    )


def save_generation_memory_to_github_st(memory: dict, sha: str | None) -> str | None:
    g = _github_cfg()
    resp = gmem.save_memory_to_github(
        memory,
        owner=g["owner"],
        repo=g["repo"],
        token=g["token"],
        branch=g.get("branch", "main"),
        sha=sha,
        path=_generation_memory_path(),
    )
    try:
        content = resp.get("content") if isinstance(resp, dict) else None
        if isinstance(content, dict) and content.get("sha"):
            return str(content["sha"])
    except Exception:
        pass
    return None


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
    rows, sha = usvc.load_doctor_rows(_io_config(), doctor)
    _official_rows_cache()[doctor] = (rows, sha)
    return rows, sha


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
    """Return True if the HTTP error indicates a concurrent update (SHA mismatch / conflict)."""
    return usvc.is_conflict_error(err)


# ---------------- Doctor slug ----------------
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
            _pin_record_cached.clear()
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

@st.cache_data(ttl=300, show_spinner=False)
def _pin_record_cached(doctor: str) -> dict | None:
    """PIN record for login checks (the login page re-reads it at every rerun)."""
    return load_doctor_pin_record(doctor)[0]


def verify_doctor_pin(doctor: str, pin: str) -> bool:
    """Verify PIN against GitHub record; fallback to secrets doctor_pins for migration."""
    rec = _pin_record_cached(doctor)
    if isinstance(rec, dict) and rec.get("hash_b64") and rec.get("salt_b64"):
        return _verify_pin(pin, rec)
    # migration fallback (old secrets-based PINs)
    pins = _get_doctor_pins()
    expected = str(pins.get(doctor, ""))
    return bool(pin) and bool(expected) and (pin == expected)

def doctor_has_pin(doctor: str) -> bool:
    rec = _pin_record_cached(doctor)
    return bool(isinstance(rec, dict) and rec.get("hash_b64") and rec.get("salt_b64"))


def _send_otp_email(dest: str, code: str) -> None:
    _send_email(
        [dest],
        [],
        "Codice verifica – Turni UTIC",
        "Hai richiesto un codice per impostare o resettare il PIN di accesso.\n\n"
        f"CODICE: {code}\n\n"
        "Se non sei stato tu, ignora questo messaggio.",
    )


def _send_email(to: list[str], cc: list[str], subject: str, body: str) -> None:
    usvc.send_email(_smtp_cfg(), to, cc, subject, body)

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


def _parse_utc_iso(ts: str) -> datetime | None:
    s = str(ts or "").strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


# ---------------- Sessioni medico (lease in memoria) ----------------
# Tutte le sessioni Streamlit dell'app vivono in questo processo: il lease
# "chi sta modificando" sta in memoria, senza un commit GitHub a ogni login e
# heartbeat (erano metà dei commit) né una lettura ogni 5 secondi.
@st.cache_resource
def _lease_registry() -> session_leases.LeaseRegistry:
    return session_leases.LeaseRegistry(ttl_seconds=DOCTOR_SESSION_TTL_MINUTES * 60)


def acquire_doctor_session_lease(*, doctor: str, session_id: str) -> None:
    """New login: this session takes over (older sessions get kicked out)."""
    _lease_registry().acquire((doctor or "").strip(), session_id)


def check_doctor_session_lease(doctor: str, session_id: str) -> bool:
    """True if this session owns the lease, or nobody holds a live one."""
    return _lease_registry().is_active((doctor or "").strip(), session_id)


def touch_doctor_session_lease(doctor: str, session_id: str) -> bool:
    """Heartbeat: keep our lease alive; never steal a newer session's lease."""
    return _lease_registry().touch((doctor or "").strip(), session_id)


def _month_entries_signature(rows: list[dict]) -> list[tuple[str, str, str]]:
    """Signature of one doctor-month rows: sorted (date, shift, note)."""
    return ustore.rows_signature(rows)


def _entries_signature_from_tuples(entries: list[tuple[date, str, str]]) -> list[tuple[str, str, str]]:
    """Signature of editor entries (date, shift, note)."""
    return ustore.entries_signature(entries)


# ---------------- Servizio salvataggi (GitHub + coda + mail) ----------------
def _receipt_cc_list(settings: dict | None) -> list[str]:
    cc = list((settings or {}).get("receipt_cc_emails") or [])
    extra = _get_secret(("notifications", "receipt_cc"), None)
    if isinstance(extra, str):
        cc += receipts.parse_email_list(extra)
    elif isinstance(extra, (list, tuple)):
        cc += [str(x) for x in extra]
    return cc


def _io_config() -> usvc.ServiceConfig:
    """GitHub-only configuration (no settings read)."""
    return usvc.ServiceConfig(github=dict(_github_cfg()), app_build=APP_BUILD)


def _service_config() -> usvc.ServiceConfig:
    """Full configuration captured in the page thread (the worker never reads st.*)."""
    try:
        settings, _ = load_app_settings_from_github()
    except Exception:
        settings = dict(DEFAULT_SETTINGS)
    return usvc.ServiceConfig(
        github=dict(_github_cfg()),
        smtp=dict(_smtp_cfg()),
        receipt_cc=_receipt_cc_list(settings),
        app_build=APP_BUILD,
    )


@st.cache_resource
def _official_rows_cache() -> dict:
    """Last official rows read/written per doctor: fallback when GitHub is saturated."""
    return {}


# ---------------- Bozze indisponibilità (autosalvataggio) ----------------
# Le modifiche non ancora inviate con "Salva" finiscono subito in una bozza in
# memoria (nessuna chiamata GitHub) e una copia su GitHub al massimo ogni 2 minuti
# (sopravvive ai riavvii, fa da prova). Il solver usa SOLO il CSV ufficiale.
DRAFT_PERSIST_SECONDS = 120         # copia della bozza su GitHub al massimo ogni 2 minuti
DRAFT_AUTOSAVE_FLUSH_SECONDS = 20   # controllo periodico (banner + copia in sospeso)


@st.cache_resource
def _draft_store() -> udrafts.MemoryDraftStore:
    cfg = _io_config()
    return udrafts.MemoryDraftStore(
        load_fn=lambda d: usvc.load_draft(cfg, d, max_wait=5),
        persist_fn=lambda d, fn: usvc.update_draft(cfg, d, fn, max_wait=5),
        persist_min_interval=DRAFT_PERSIST_SECONDS,
    )


def _queue_worker_interval() -> float:
    try:
        return float(os.environ.get("TURNI_QUEUE_WORKER_INTERVAL") or 10.0)
    except ValueError:
        return 10.0


@st.cache_resource
def _save_queue() -> usvc.SaveQueue:
    """Saves waiting for GitHub (rate limit / outage), completed by a worker thread."""
    store = _draft_store()
    rows_cache = _official_rows_cache()

    def _on_saved(result: usvc.SaveResult) -> None:  # worker thread: no st.* here
        for (yy, mm) in result.job.entries_by_month:
            store.remove_month(result.job.doctor, f"{int(yy):04d}-{int(mm):02d}")
        store.flush(result.job.doctor, force=True)
        rows_cache[result.job.doctor] = (result.outcome.verified_rows, result.outcome.sha)

    q = usvc.SaveQueue(
        _service_config(),
        persist_path=Path(tempfile.gettempdir()) / "turni_unav_save_queue.json",
        on_saved=_on_saved,
    )
    q.start_worker(interval=_queue_worker_interval())
    return q


def _draft_state_key(doctor: str) -> str:
    return f"unav_draft_state::{doctor}"


def _draft_state(doctor: str) -> dict:
    state = st.session_state.get(_draft_state_key(doctor))
    if not isinstance(state, dict):
        state = {"written": {}, "pending": {}, "unsent": {}, "last_write_at": "", "error": ""}
        st.session_state[_draft_state_key(doctor)] = state
    return state


def _session_draft(doctor: str) -> dict:
    """Current draft of the doctor (memory, loaded from GitHub the first time)."""
    return _draft_store().get(doctor)


def _draft_mark_written(doctor: str, mk: str, entries) -> None:
    state = _draft_state(doctor)
    state["written"][mk] = ustore.entries_signature(entries)
    state["pending"].pop(mk, None)


def _draft_note_change(doctor: str, mk: str, entries, base_signature, session_id: str) -> None:
    state = _draft_state(doctor)
    sig = ustore.entries_signature(entries)
    if sig == state["written"].get(mk):
        state["pending"].pop(mk, None)
        return
    state["pending"][mk] = {
        "entries": [(d, sh, note) for d, sh, note in entries],
        "base_signature": list(base_signature or []),
        "session_id": session_id,
    }


def _draft_flush(doctor: str, *, force: bool = False) -> None:
    """Pending edits → memory draft (always); memory → GitHub copy (throttled)."""
    state = _draft_state(doctor)
    store = _draft_store()
    pending = dict(state.get("pending") or {})
    if pending:
        now_iso = _utc_now_iso()
        for mk, item in pending.items():
            store.set_month(
                doctor,
                mk,
                entries=item["entries"],
                base_signature=item["base_signature"],
                updated_at=now_iso,
                session_id=item["session_id"],
            )
            state["written"][mk] = ustore.entries_signature(item["entries"])
            if state["pending"].get(mk) is item:
                state["pending"].pop(mk, None)
        state["last_write_at"] = now_iso
    copied = store.flush(doctor, force=force)
    state["error"] = "" if copied or not force else "copia su GitHub rimandata (GitHub saturo)"


def _draft_discard_month(doctor: str, mk: str) -> None:
    store = _draft_store()
    store.remove_month(doctor, mk)
    store.flush(doctor, force=True)


def _local_hhmm(ts_utc: str) -> str:
    try:
        t = datetime.fromisoformat(str(ts_utc).replace("Z", "+00:00"))
        return t.astimezone(receipts.LOCAL_TZ).strftime("%d/%m %H:%M")
    except Exception:
        return str(ts_utc or "")


@st.fragment(run_every=DRAFT_AUTOSAVE_FLUSH_SECONDS)
def _draft_status_and_flush(doctor: str, session_id: str, mk: str) -> None:
    """Periodic draft copy + unsent-changes banner for month mk."""
    state = _draft_state(doctor)
    if check_doctor_session_lease(doctor, session_id):
        _draft_flush(doctor)
    unsent = (state.get("unsent") or {}).get(mk)
    if unsent:
        in_draft = mk not in (state.get("pending") or {}) and bool(state.get("last_write_at"))
        where = (
            f"Sono al sicuro nella bozza (salvata alle {_local_hhmm(state['last_write_at'])}) anche se chiudi la pagina, ma "
            if in_draft
            else "Salvataggio in bozza in corso… "
        )
        st.warning(
            f"⚠️ **Modifiche NON inviate** (+{unsent['added']} / -{unsent['removed']} rispetto a quanto registrato). "
            f"{where}**non verranno usate per i turni finché non premi 💾 Salva indisponibilità.**"
        )
    else:
        st.caption("✅ Tutto inviato: quello che vedi è esattamente quanto registrato sul server.")
    if state.get("error"):
        st.caption(f"ℹ️ {state['error']}: riprovo automaticamente.")


def _reset_doctor_editor_state(doctor: str) -> None:
    """Forget editor rows/drafts held in this browser session.

    After a login or a reload the editor must restart from the server data (and
    the draft), never from a stale in-memory copy of another session.
    """
    prefixes = (
        f"unav_rows_{doctor}_",
        f"avail_rows_{doctor}_",
        f"avail_store_baseline_{doctor}",
        _draft_state_key(doctor),
        f"unav_outbox::{doctor}",
    )
    for k in list(st.session_state.keys()):
        if str(k).startswith(prefixes):
            st.session_state.pop(k, None)


@st.cache_data(ttl=120, show_spinner=False)
def _github_drafts_for_months(month_keys: tuple[str, ...]) -> dict:
    """Draft files on GitHub touching the months: {doctor: draft}."""
    cfg = _io_config()
    g = cfg.github
    files = github_utils.list_dir(
        g["owner"], g["repo"], str(g.get("drafts_dir") or udrafts.DRAFTS_DIR_DEFAULT), g["token"],
        g.get("branch", "main"), max_wait=10,
    )
    out: dict = {}
    for f in files:
        if not str(f.get("name", "")).endswith(".json"):
            continue
        gf = github_utils.get_file(g["owner"], g["repo"], f["path"], g["token"], g.get("branch", "main"), max_wait=10)
        if gf is None:
            continue
        draft = udrafts.from_text(gf.text)
        if draft.get("doctor") and any(mk in (draft.get("months") or {}) for mk in month_keys):
            out[draft["doctor"]] = draft
    return out


@st.cache_data(ttl=30, show_spinner=False)
def load_pending_drafts_summary(month_keys: tuple[str, ...]) -> list[dict]:
    """Unsent drafts (draft != official data) for the given months, all doctors."""
    drafts = dict(_github_drafts_for_months(month_keys))
    drafts.update(_draft_store().all_drafts())  # memory is the most recent copy
    out: list[dict] = []
    for doctor, draft in sorted(drafts.items()):
        if not any(mk in (draft.get("months") or {}) for mk in month_keys):
            continue
        official, _ = load_doctor_unavail_from_github(doctor)
        out += udrafts.pending_summary(draft, official, month_keys)
    return out


def _apply_saved_result_to_session(doctor: str, result: usvc.SaveResult, selected) -> None:
    """Server data just verified by read-back → baseline + editor bases (no reload)."""
    _official_rows_cache()[doctor] = (result.outcome.verified_rows, result.outcome.sha)
    sel = tuple(tuple(x) for x in (selected or []))
    cur = st.session_state.get(_BASELINE_SS_KEY)
    if isinstance(cur, dict) and cur.get("doctor") == doctor:
        sel = tuple(cur.get("selected") or sel)
    _set_doctor_baseline_rows(doctor, sel, result.outcome.verified_rows, result.outcome.sha)
    state = _draft_state(doctor)
    for (yy, mm), entries in result.job.entries_by_month.items():
        mk = f"{int(yy):04d}-{int(mm):02d}"
        st.session_state[f"unav_rows_{doctor}_{yy}_{mm}__base_sig"] = ustore.month_signature(
            result.outcome.verified_rows, doctor, int(yy), int(mm)
        )
        _draft_mark_written(doctor, mk, entries)
        state["unsent"].pop(mk, None)


def _saved_message(result: usvc.SaveResult, mail_status: str, warnings: list[str]) -> str:
    outcome = result.outcome
    if outcome.changed:
        parts = [
            f"{receipts.month_label(mk)}: +{d.get('added_count', 0)} / -{d.get('removed_count', 0)}, "
            f"ora registrate {d.get('after_count', 0)}"
            for mk, d in sorted(outcome.diffs.items())
        ]
        msg = "✅ Salvataggio completato e verificato sul server — " + "; ".join(parts) + "."
        if mail_status:
            msg += f" {mail_status}"
    else:
        msg = "✅ Nessuna modifica da inviare: il server ha già esattamente queste indisponibilità."
    if warnings:
        msg += " " + " ".join(f"⚠️ {w}" for w in warnings)
    return msg


def _registered_details(doctor: str, result: usvc.SaveResult) -> str:
    rows_by_month = {
        f"{int(yy):04d}-{int(mm):02d}": ustore.filter_doctor_month(result.outcome.verified_rows, doctor, int(yy), int(mm))
        for (yy, mm) in result.job.entries_by_month
    }
    return receipts.build_receipt(doctor, rows_by_month, {}, saved_at_utc=result.saved_at, commit_sha=result.outcome.commit_sha)[1]


def _process_unav_outbox(doctor: str, selected) -> None:
    """Complete a successful save: baseline, draft cleanup, audit, receipt, message.

    Every step is recorded in the outbox as soon as it is done and no st.* element
    is rendered before the end: if a double tap interrupts the run right after
    the GitHub write, the next run finishes the job without repeating steps.
    """
    key = f"unav_outbox::{doctor}"
    box = st.session_state.get(key)
    if not isinstance(box, dict):
        return
    result: usvc.SaveResult = box["result"]
    cfg: usvc.ServiceConfig = box["cfg"]
    done = box.setdefault("done", set())

    if "baseline" not in done:
        _apply_saved_result_to_session(doctor, result, box.get("selected") or selected)
        done.add("baseline")
    if "draft" not in done:
        store = _draft_store()
        for (yy, mm) in result.job.entries_by_month:
            store.remove_month(doctor, f"{int(yy):04d}-{int(mm):02d}")
        store.flush(doctor, force=True)
        done.add("draft")
    warnings = box.setdefault("warnings", [])
    if result.outcome.changed and "audit" not in done:
        warnings += usvc.write_audit(cfg, result)
        done.add("audit")
    if result.outcome.changed and "mail" not in done:
        box["mail_status"] = usvc.send_receipt(cfg, result)
        done.add("mail")

    set_unav_flash(
        doctor, "success", _saved_message(result, box.get("mail_status", ""), warnings),
        details=_registered_details(doctor, result),
    )
    st.session_state.pop(key, None)


def _show_save_queue_status(doctor: str, selected) -> None:
    """Results of saves completed in background + saves still waiting for GitHub."""
    q = _save_queue()
    for res in q.take_results(doctor):
        result = res.get("result")
        if res["status"] == "done" and result is not None:
            _apply_saved_result_to_session(doctor, result, selected)
            set_unav_flash(doctor, "success", res["message"], details=_registered_details(doctor, result))
        else:
            set_unav_flash(doctor, "error", res["message"])
    for job in q.pending(doctor):
        months = ", ".join(receipts.month_label(f"{int(yy):04d}-{int(mm):02d}") for (yy, mm) in sorted(job.entries_by_month))
        when = datetime.fromtimestamp(job.created_at, receipts.LOCAL_TZ).strftime("%d/%m %H:%M")
        st.warning(
            f"⏳ **Salvataggio IN CODA** ({months}, inviato alle {when}): GitHub è momentaneamente saturo. "
            "Verrà registrato automaticamente appena possibile e riceverai la mail di conferma. "
            "Non serve premere di nuovo Salva. **Finché non arriva la mail non è registrato.**"
        )


# ---------------- GitHub settings & audit log ----------------
DEFAULT_SETTINGS = {
    "unavailability_open": True,
    "max_unavailability_per_shift": 6,
    "max_availability_per_shift": 6,  # max preferenze disponibilità per fascia per mese
    "max_weekend_days": MAX_WEEKEND_DAYS,  # max sabati e max domeniche distinti per mese
    "doctor_caps": {},  # cap personalizzato per medico: {"Dattilo": 10, "De Gregorio": 10, "Zito": 10}
    # Copie della mail di resoconto a ogni salvataggio (oltre al medico).
    # Modificabile dal pannello admin; si somma a secrets [notifications] receipt_cc.
    "receipt_cc_emails": ["utic@polime.it"],
}

AUDIT_FIELDS = usvc.AUDIT_FIELDS

@st.cache_data(ttl=60, show_spinner=False)
def load_app_settings_from_github() -> tuple[dict, str | None]:
    """App settings, cached 60 s (read at every doctor rerun). Cleared on save."""
    return _load_app_settings_uncached()


def _load_app_settings_uncached() -> tuple[dict, str | None]:
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

    cc = data.get("receipt_cc_emails")
    if isinstance(cc, str):
        cc = receipts.parse_email_list(cc)
    if isinstance(cc, list):
        out["receipt_cc_emails"] = [str(x).strip() for x in cc if str(x).strip()]

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
    load_app_settings_from_github.clear()


@st.cache_data(ttl=120, show_spinner=False)
def _pool_config_for_doctor_list() -> dict:
    """Pool config used only to list doctors on every page rerun (cached)."""
    return load_pool_config_from_github_st()[0]


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


def sync_pool_contacts_to_github(pool_cfg: dict) -> tuple[bool, str]:
    """Aggiorna doctor_contacts.yml su GitHub con le email definite nel pool_config.

    Per ogni medico attivo con email valorizzata nel pool_config, aggiunge/aggiorna
    la voce corrispondente in doctor_contacts.yml. Non cancella voci esistenti.
    Ritorna (ok, messaggio).
    """
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo")):
        return False, "GitHub non configurato — impossibile sincronizzare i contatti."

    contacts_path = g.get("contacts_path") or "data/doctor_contacts.yml"

    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"], path=contacts_path,
        token=g["token"], branch=g.get("branch", "main"),
    )
    existing: dict = {}
    sha_contacts: str | None = None
    if gf:
        try:
            existing = yaml.safe_load(gf.text) or {}
        except Exception:
            existing = {}
        sha_contacts = gf.sha

    changed = False
    for doc, dcfg in (pool_cfg.get("doctors") or {}).items():
        if not dcfg.get("active", True):
            continue
        email = str(dcfg.get("email") or "").strip()
        if not email:
            continue
        if not isinstance(existing.get(doc), dict):
            existing[doc] = {}
        if existing[doc].get("email") != email:
            existing[doc]["email"] = email
            changed = True

    if not changed:
        return True, "Nessuna email nuova — doctor_contacts.yml già aggiornato."

    content = yaml.dump(
        {k: existing[k] for k in sorted(existing)},
        allow_unicode=True, default_flow_style=False,
    )
    github_utils.put_file(
        owner=g["owner"], repo=g["repo"], path=contacts_path,
        token=g["token"],
        message="Aggiornamento contatti medici da pool_config GUI",
        text=content,
        branch=g.get("branch", "main"),
        sha=sha_contacts,
    )
    load_doctor_contacts_from_github.clear()  # invalida cache
    return True, "Email medici sincronizzate con doctor_contacts.yml ✓"


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
            _pool_config_for_doctor_list.clear()
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

def compute_unavailability_diff(existing_rows: list[dict], new_entries: list[tuple[date, str, str]]) -> dict:
    """Return a diff summary between current store rows and the edited entries."""
    return ustore.compute_month_diff(existing_rows, new_entries)

def append_unavailability_audit_log(mk: str, row: dict, max_retries: int = 10):
    """Append a row to the monthly audit log on GitHub (idempotent, conflict-retried)."""
    usvc.append_audit_row(_io_config(), mk, row, max_retries=max_retries)

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


def _month_bounds(yy: int, mm: int) -> tuple[date, date]:
    first_day = date(int(yy), int(mm), 1)
    if int(mm) == 12:
        last_day = date(int(yy) + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(int(yy), int(mm) + 1, 1) - timedelta(days=1)
    return first_day, last_day


def _store_rows_to_editor_rows(rows: list[dict], yy: int, mm: int, *, availability: bool = False) -> list[dict]:
    out = []
    for r in rows:
        try:
            d = date.fromisoformat(str(r.get("date") or "")[:10])
        except Exception:
            d = date(int(yy), int(mm), 1)
        item = {
            "id": str(uuid.uuid4()),
            "Data": d,
            "Fascia": r.get("shift", "Mattina"),
            "Note": r.get("note", ""),
        }
        if availability:
            item["Priorita"] = ustore.norm_priority(r.get("priority", "media"))
        out.append(item)
    return out


def _dedup_unav_editor_rows(rows: list[dict]) -> list[tuple[date, str, str]]:
    dedup: dict[tuple[date, str], str] = {}
    for r in rows or []:
        d = r.get("Data")
        if isinstance(d, datetime):
            d = d.date()
        if not isinstance(d, date):
            continue
        sh, _changed, _unknown = normalize_fascia(r.get("Fascia", ""))
        if not sh:
            continue
        dedup[(d, sh)] = str(r.get("Note", "") or "")
    return [(d, sh, note) for (d, sh), note in sorted(dedup.items(), key=lambda kv: (kv[0][0], kv[0][1]))]


def _dedup_avail_editor_rows(rows: list[dict]) -> list[tuple[date, str, str, str]]:
    dedup: dict[tuple[date, str], tuple[str, str]] = {}
    for r in rows or []:
        d = r.get("Data")
        if isinstance(d, datetime):
            d = d.date()
        if not isinstance(d, date):
            continue
        sh, _changed, _unknown = normalize_fascia(r.get("Fascia", ""))
        if not sh or sh == "Ferie":
            continue
        pri = ustore.norm_priority(r.get("Priorita", "media"))
        dedup[(d, sh)] = (str(r.get("Note", "") or ""), pri)
    return [
        (d, sh, note, pri)
        for (d, sh), (note, pri) in sorted(dedup.items(), key=lambda kv: (kv[0][0], kv[0][1]))
    ]


def _entries_summary(entries: list[tuple]) -> str:
    if not entries:
        return "0 righe"
    dates = sorted(e[0] for e in entries if e and isinstance(e[0], date))
    counts: dict[str, int] = {}
    for e in entries:
        if len(e) >= 2:
            sh = str(e[1] or "")
            counts[sh] = counts.get(sh, 0) + 1
    parts = [f"{len(entries)} righe"]
    if dates:
        parts.append(f"{dates[0]:%d/%m/%Y}-{dates[-1]:%d/%m/%Y}")
    if counts:
        parts.append(", ".join(f"{k}: {v}" for k, v in sorted(counts.items())))
    return " | ".join(parts)


def _rerun_fragment_or_app() -> None:
    try:
        st.rerun(scope="fragment")
    except Exception:
        st.rerun()


def _render_admin_month_rows(
    *,
    rows_key: str,
    rows: list[dict],
    first_day: date,
    last_day: date,
    fascia_options: list[str],
    availability: bool = False,
) -> list[dict]:
    if rows_key not in st.session_state:
        st.session_state[rows_key] = rows or []

    cur_rows = list(st.session_state.get(rows_key) or [])
    if not cur_rows:
        st.caption("Nessuna riga: usa ➕ Aggiungi riga o Aggiungi periodo.")

    with st.expander("Aggiungi periodo", expanded=False):
        st.caption("Aggiunge righe alla lista qui sotto. Per scriverle su GitHub devi poi premere Salva.")
        with st.form(f"{rows_key}_period_form", clear_on_submit=False):
            p1, p2, p3, p4 = st.columns([1.2, 1.2, 1.4, 1], vertical_alignment="bottom")
            with p1:
                p_start = st.date_input("Dal", value=first_day, min_value=first_day, max_value=last_day, key=f"{rows_key}_period_start", format="DD/MM/YYYY")
            with p2:
                p_end = st.date_input("Al", value=first_day, min_value=first_day, max_value=last_day, key=f"{rows_key}_period_end", format="DD/MM/YYYY")
            with p3:
                default_shift = "Ferie" if (not availability and "Ferie" in fascia_options) else "Mattina"
                p_shift = st.selectbox(
                    "Fascia",
                    fascia_options,
                    index=fascia_options.index(default_shift) if default_shift in fascia_options else 0,
                    key=f"{rows_key}_period_shift",
                )
            p_note = st.text_input("Note periodo", key=f"{rows_key}_period_note")
            p_priority = "media"
            if availability:
                p_priority = st.selectbox("Priorità periodo", ["alta", "media", "bassa"], index=1, key=f"{rows_key}_period_priority")
            with p4:
                add_period = st.form_submit_button("Aggiungi alla lista", use_container_width=True)

        if add_period:
            if p_end < p_start:
                st.error("La data finale deve essere uguale o successiva alla data iniziale.")
            else:
                existing = {(r.get("Data"), r.get("Fascia")) for r in cur_rows}
                for offset in range((p_end - p_start).days + 1):
                    d = p_start + timedelta(days=offset)
                    if (d, p_shift) in existing:
                        continue
                    cur_rows.append({
                        "id": str(uuid.uuid4()),
                        "Data": d,
                        "Fascia": p_shift,
                        "Note": p_note,
                        **({"Priorita": p_priority} if availability else {}),
                    })
                st.session_state[rows_key] = cur_rows
                _rerun_fragment_or_app()

    remove_ids: set[str] = set()
    updated: list[dict] = []
    for r in cur_rows:
        rid = str(r.get("id") or uuid.uuid4())
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2 if availability else 0.05, 0.45], vertical_alignment="bottom")
        with c1:
            d_val = st.date_input(
                "Data",
                value=r.get("Data") if isinstance(r.get("Data"), date) else first_day,
                min_value=first_day,
                max_value=last_day,
                key=f"{rows_key}_{rid}_date",
                format="DD/MM/YYYY",
            )
        with c2:
            prev_shift = r.get("Fascia", "Mattina")
            sh_val = st.selectbox(
                "Fascia",
                fascia_options,
                index=fascia_options.index(prev_shift) if prev_shift in fascia_options else 0,
                key=f"{rows_key}_{rid}_shift",
            )
        pri_val = "media"
        if availability:
            with c3:
                prev_pri = ustore.norm_priority(r.get("Priorita", "media"))
                pri_val = st.selectbox(
                    "Priorità",
                    ["alta", "media", "bassa"],
                    index=["alta", "media", "bassa"].index(prev_pri),
                    key=f"{rows_key}_{rid}_priority",
                )
        with c4:
            if st.button("🗑", key=f"{rows_key}_{rid}_remove", help="Rimuovi riga"):
                remove_ids.add(rid)
        note_val = st.text_input("Note", value=str(r.get("Note", "") or ""), key=f"{rows_key}_{rid}_note")
        if rid not in remove_ids:
            item = {"id": rid, "Data": d_val, "Fascia": sh_val, "Note": note_val}
            if availability:
                item["Priorita"] = pri_val
            updated.append(item)

    if remove_ids:
        st.session_state[rows_key] = updated
        _rerun_fragment_or_app()

    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("➕ Aggiungi riga", key=f"{rows_key}_add_row", use_container_width=True):
            updated.append({
                "id": str(uuid.uuid4()),
                "Data": first_day,
                "Fascia": "Mattina",
                "Note": "",
                **({"Priorita": "media"} if availability else {}),
            })
            st.session_state[rows_key] = updated
            _rerun_fragment_or_app()
    with b2:
        if st.button("🧹 Pulisci", key=f"{rows_key}_clean", use_container_width=True):
            st.session_state[rows_key] = [
                r for r in updated
                if str(r.get("Note", "") or "").strip() or r.get("Data") != first_day or r.get("Fascia") != "Mattina"
            ]
            _rerun_fragment_or_app()

    st.session_state[rows_key] = updated
    return updated


@st.fragment
def render_admin_doctor_data_editor(doctors: list[str], default_year: int, default_month: int) -> None:
    st.caption("Modifica direttamente i file per-medico usati poi dalla generazione.")
    if not doctors:
        st.info("Nessun medico configurato.")
        return

    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        doctor = st.selectbox("Medico", doctors, key="admin_store_editor_doctor")
    with c2:
        yy = st.number_input("Anno", min_value=2025, max_value=2035, value=int(default_year), step=1, key="admin_store_editor_year")
    with c3:
        mm = st.number_input("Mese", min_value=1, max_value=12, value=int(default_month), step=1, key="admin_store_editor_month")

    first_day, last_day = _month_bounds(int(yy), int(mm))
    reload_key = f"admin_store_reload_{doctor}_{int(yy)}_{int(mm)}"
    if st.button("Ricarica dati medico", key=f"{reload_key}_button"):
        for prefix in ("admin_unav_rows", "admin_avail_rows"):
            st.session_state.pop(f"{prefix}_{doctor}_{int(yy)}_{int(mm)}", None)
            st.session_state.pop(f"{prefix}_{doctor}_{int(yy)}_{int(mm)}__base_sig", None)
        _rerun_fragment_or_app()

    try:
        unav_rows, unav_sha = load_doctor_unavail_from_github(doctor)
    except Exception as e:
        st.error(f"Errore lettura indisponibilità: {e}")
        unav_rows, unav_sha = [], None
    try:
        avail_rows, avail_sha = load_doctor_avail_from_github(doctor)
    except Exception as e:
        st.error(f"Errore lettura preferenze: {e}")
        avail_rows, avail_sha = [], None

    tab_unav, tab_avail = st.tabs(["Indisponibilità", "Preferenze"])
    flash_key = f"admin_store__{doctor}__{int(yy)}__{int(mm)}"
    render_unav_flash(flash_key)

    with tab_unav:
        existing = ustore.filter_doctor_month(unav_rows, doctor, int(yy), int(mm))
        existing_entries = [
            (
                ustore.parse_iso_date(r.get("date", "")),
                r.get("shift", ""),
                r.get("note", ""),
            )
            for r in existing
            if r.get("date") and r.get("shift")
        ]
        st.caption("Persistito su GitHub: " + _entries_summary(existing_entries))
        rows_key = f"admin_unav_rows_{doctor}_{int(yy)}_{int(mm)}"
        base_sig_key = f"{rows_key}__base_sig"
        if rows_key not in st.session_state:
            st.session_state[base_sig_key] = _month_entries_signature(existing)
        editor_rows = _render_admin_month_rows(
            rows_key=rows_key,
            rows=_store_rows_to_editor_rows(existing, int(yy), int(mm), availability=False),
            first_day=first_day,
            last_day=last_day,
            fascia_options=FASCIA_OPTIONS,
            availability=False,
        )
        entries = _dedup_unav_editor_rows(editor_rows)
        st.caption(f"Totale righe salvabili: {len(entries)}")
        if _entries_signature_from_tuples(entries) != _month_entries_signature(existing):
            st.warning("Ci sono modifiche nella lista non ancora salvate su GitHub.", icon="⚠️")
        if st.button("💾 Salva indisponibilità medico", key=f"{rows_key}_save", type="primary"):
            admin_cfg = _service_config()
            admin_job = usvc.make_job(
                doctor=doctor,
                doctor_email=str(get_doctor_contact(doctor).get("email") or ""),
                entries_by_month={(int(yy), int(mm)): entries},
                base_signatures={
                    (int(yy), int(mm)): st.session_state.get(base_sig_key, _month_entries_signature(existing))
                },
                by_admin=True,
                origin=f"admin:{rows_key}",
            )
            try:
                queue = _save_queue()
                queue.cfg = admin_cfg
                status, result = queue.save_or_enqueue(admin_job, max_wait=25)
                if status == "queued":
                    set_unav_flash(
                        flash_key,
                        "warning",
                        f"⏳ GitHub saturo: salvataggio per {doctor} IN CODA, verrà registrato automaticamente "
                        "(il medico riceverà la mail a registrazione avvenuta).",
                    )
                else:
                    warnings = usvc.write_audit(admin_cfg, result) if result.outcome.changed else []
                    mail_status = usvc.send_receipt(admin_cfg, result) if result.outcome.changed else ""
                    set_unav_flash(
                        flash_key,
                        "success",
                        (
                            f"Indisponibilità salvate e verificate su GitHub per {doctor}. {mail_status}"
                            if result.outcome.changed
                            else f"Nessuna modifica da salvare per {doctor}."
                        )
                        + "".join(f" ⚠️ {w}" for w in warnings),
                        details=_entries_summary(entries),
                    )
                st.session_state.pop(rows_key, None)
                st.session_state.pop(base_sig_key, None)
                _rerun_fragment_or_app()
            except ustore.MonthConflictError as e:
                set_unav_flash(
                    flash_key,
                    "error",
                    f"Salvataggio bloccato per {doctor}: {e} Premi 'Ricarica dati medico' e ripeti la modifica.",
                )
                _rerun_fragment_or_app()
            except Exception as e:
                set_unav_flash(
                    flash_key,
                    "error",
                    f"Errore salvataggio indisponibilità per {doctor}: {e}",
                )
                _rerun_fragment_or_app()

    with tab_avail:
        existing = ustore.filter_doctor_month(avail_rows, doctor, int(yy), int(mm))
        existing_entries = [
            (
                ustore.parse_iso_date(r.get("date", "")),
                r.get("shift", ""),
                r.get("note", ""),
                r.get("priority", "media"),
            )
            for r in existing
            if r.get("date") and r.get("shift")
        ]
        st.caption("Preferenze persistite su GitHub: " + _entries_summary(existing_entries))
        rows_key = f"admin_avail_rows_{doctor}_{int(yy)}_{int(mm)}"
        editor_rows = _render_admin_month_rows(
            rows_key=rows_key,
            rows=_store_rows_to_editor_rows(existing, int(yy), int(mm), availability=True),
            first_day=first_day,
            last_day=last_day,
            fascia_options=AVAIL_FASCIA_OPTIONS,
            availability=True,
        )
        entries = _dedup_avail_editor_rows(editor_rows)
        st.caption(f"Totale righe salvabili: {len(entries)}")
        if _entries_signature_from_tuples([(d, sh, note) for d, sh, note, _pri in entries]) != _month_entries_signature(existing):
            st.warning("Ci sono preferenze nella lista non ancora salvate su GitHub.", icon="⚠️")
        if st.button("💾 Salva preferenze medico", key=f"{rows_key}_save", type="primary"):
            updated_at = _utc_now_iso()
            try:
                save_doctor_availability_with_retry(
                    doctor=doctor,
                    entries_by_month={(int(yy), int(mm)): entries},
                    updated_at=updated_at,
                    message=f"Admin update availability: {doctor} ({updated_at})",
                    initial_rows=avail_rows,
                    initial_sha=avail_sha,
                )
                set_unav_flash(
                    flash_key,
                    "success",
                    f"Preferenze salvate e verificate su GitHub per {doctor}.",
                    details=_entries_summary(entries),
                )
                st.session_state.pop(rows_key, None)
                _rerun_fragment_or_app()
            except Exception as e:
                set_unav_flash(
                    flash_key,
                    "error",
                    f"Errore salvataggio preferenze per {doctor}: {e}",
                )
                _rerun_fragment_or_app()


# ---------------- Medico UX: baseline snapshot + session guard ----------------
_BASELINE_SS_KEY = "unav_store_baseline"


def get_or_load_doctor_baseline(
    doctor: str,
    selected_months: list[tuple[int, int]],
    force_reload: bool = False,
) -> dict:
    """Return the server snapshot the doctor's editors start from.

    Each month editor also keeps the signature it started from
    (`<rows_key>__base_sig`): that is what the save compares to the server to
    detect a month changed by another session.
    """
    doctor = (doctor or "").strip()
    selected_key = tuple((int(y), int(m)) for (y, m) in (selected_months or []))

    cur = st.session_state.get(_BASELINE_SS_KEY)
    if (
        (not force_reload)
        and isinstance(cur, dict)
        and cur.get("doctor") == doctor
        and tuple(cur.get("selected") or ()) == selected_key
    ):
        return cur

    # Load only this doctor's file — per-doctor storage, no cross-doctor SHA conflict.
    try:
        rows, sha = load_doctor_unavail_from_github(doctor)
    except github_utils.GithubUnavailable:
        cached = _official_rows_cache().get(doctor)
        if cached is None:
            raise
        # GitHub saturo: si lavora sugli ultimi dati letti. Il salvataggio ricontrolla
        # comunque il server (o va in coda) e non sovrascrive mai modifiche altrui.
        rows, sha = cached
        new = _set_doctor_baseline_rows(doctor, selected_key, rows, sha)
        new["stale"] = True
        return new
    return _set_doctor_baseline_rows(doctor, selected_key, rows, sha)


def _set_doctor_baseline_rows(doctor: str, selected_key, rows: list[dict], sha: str | None) -> dict:
    new = {
        "doctor": (doctor or "").strip(),
        "selected": tuple(selected_key),
        "rows": rows,
        "sha": sha,
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
    """Single-session guard per doctor (in-memory lease, no GitHub calls).

    - First entry after login: takes the lease (kicking out other sessions).
    - Every rerun: heartbeat; if a newer login took the lease → forced logout.
    """
    doctor = (doctor or "").strip()
    ss_key = _doctor_session_state_key(doctor)
    cur = st.session_state.get(ss_key)
    if not isinstance(cur, dict) or not cur.get("session_id"):
        cur = {"session_id": str(uuid.uuid4()), "lease_acquired": False}
        st.session_state[ss_key] = cur
    session_id = str(cur["session_id"])
    if not cur.get("lease_acquired"):
        acquire_doctor_session_lease(doctor=doctor, session_id=session_id)
        cur["lease_acquired"] = True
        return session_id
    if not touch_doctor_session_lease(doctor, session_id):
        _logout_doctor(
            "Sessione terminata: hai effettuato accesso dallo stesso utente su un altro dispositivo/browser."
        )
        st.stop()
    return session_id


def release_doctor_session(doctor: str):
    """Logout: free the lease if this session still holds it."""
    doctor = (doctor or "").strip()
    cur = st.session_state.get(_doctor_session_state_key(doctor))
    if isinstance(cur, dict) and cur.get("session_id"):
        _lease_registry().release(doctor, str(cur["session_id"]))


@st.fragment
def render_admin_generate_panel(cfg_admin: dict, doctors: list[str], rules_path: str | Path) -> None:
    # ── Ramo Genera turni ────────────────────────────────────────────────
    # Step 1: Periodo
    st.markdown("### 1) Periodo")
    today = date.today()
    import calendar as _period_calendar

    _default_period_start = date(today.year, today.month, 1)
    _default_period_end = date(
        today.year,
        today.month,
        _period_calendar.monthrange(today.year, today.month)[1],
    )

    def _period_date_or_default(value: object, default: date) -> date:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if isinstance(value, str):
            try:
                return date.fromisoformat(value)
            except Exception:
                return default
        return default

    _period_state = st.session_state.get("admin_generation_period")
    if not isinstance(_period_state, dict):
        _period_state = {
            "mode": "Mese intero",
            "year": today.year,
            "month": today.month,
            "start": _default_period_start,
            "end": _default_period_end,
        }
        st.session_state["admin_generation_period"] = _period_state

    with st.form("admin_generation_period_form"):
        st.caption("Modifica la bozza e premi Applica periodo. Gli altri controlli non vengono ricalcolati a ogni click.")
        _current_mode = str(_period_state.get("mode") or "Mese intero")
        period_mode_draft = st.radio(
            "Tipo periodo",
            ["Mese intero", "Periodo personalizzato"],
            index=0 if _current_mode == "Mese intero" else 1,
            horizontal=True,
            key="admin_period_mode_draft",
        )

        colA, colB, colC, colD = st.columns([1, 1, 1, 1], vertical_alignment="bottom")
        with colA:
            year_draft = st.number_input(
                "Anno mese intero",
                min_value=2025,
                max_value=2035,
                value=int(_period_state.get("year") or today.year),
                step=1,
                key="admin_period_year_draft",
            )
        with colB:
            month_draft = st.number_input(
                "Mese mese intero",
                min_value=1,
                max_value=12,
                value=int(_period_state.get("month") or today.month),
                step=1,
                key="admin_period_month_draft",
            )
        with colC:
            period_start_draft = st.date_input(
                "Dal",
                value=_period_date_or_default(_period_state.get("start"), _default_period_start),
                min_value=date(2025, 1, 1),
                max_value=date(2035, 12, 31),
                format="DD/MM/YYYY",
                key="custom_period_start_draft",
            )
        with colD:
            period_end_draft = st.date_input(
                "Al",
                value=_period_date_or_default(_period_state.get("end"), _default_period_end),
                min_value=date(2025, 1, 1),
                max_value=date(2035, 12, 31),
                format="DD/MM/YYYY",
                key="custom_period_end_draft",
            )

        apply_period = st.form_submit_button("Applica periodo", type="primary")

    if apply_period:
        if period_mode_draft == "Mese intero":
            _applied_start = date(int(year_draft), int(month_draft), 1)
            _applied_end = date(
                int(year_draft),
                int(month_draft),
                _period_calendar.monthrange(int(year_draft), int(month_draft))[1],
            )
            _period_state = {
                "mode": "Mese intero",
                "year": int(year_draft),
                "month": int(month_draft),
                "start": _applied_start,
                "end": _applied_end,
            }
        else:
            if period_end_draft < period_start_draft:
                st.error("La data finale deve essere uguale o successiva alla data iniziale.")
                st.stop()
            _period_state = {
                "mode": "Periodo personalizzato",
                "year": int(period_start_draft.year),
                "month": int(period_start_draft.month),
                "start": period_start_draft,
                "end": period_end_draft,
            }
        st.session_state["admin_generation_period"] = _period_state

    period_mode = str(_period_state.get("mode") or "Mese intero")
    period_start = _period_date_or_default(_period_state.get("start"), _default_period_start)
    period_end = _period_date_or_default(_period_state.get("end"), _default_period_end)
    year = int(_period_state.get("year") or period_start.year)
    month = int(_period_state.get("month") or period_start.month)

    if period_mode == "Mese intero":
        mk = f"{int(year)}-{int(month):02d}"
        st.caption(f"Periodo attivo: **{mk}**")
    else:
        if (period_end - period_start).days > 62:
            st.warning("Periodo superiore a 63 giorni: la generazione può essere più lenta.", icon="⚠️")
        mk = f"{period_start:%Y-%m-%d}_{period_end:%Y-%m-%d}"
        st.caption(f"Periodo attivo: **{period_start:%d/%m/%Y} – {period_end:%d/%m/%Y}**")
    _period_dates = [
        period_start + timedelta(days=i)
        for i in range((period_end - period_start).days + 1)
    ]

    # Bozze NON inviate: modifiche che i medici hanno inserito ma non salvato.
    # Il solver non le usa: meglio sollecitare prima di generare.
    _period_month_keys = tuple(sorted({f"{d.year:04d}-{d.month:02d}" for d in _period_dates}))
    try:
        _pending_drafts = load_pending_drafts_summary(_period_month_keys)
    except Exception as _pd_err:
        _pending_drafts = []
        st.caption(f"Controllo bozze non inviate non riuscito: {_pd_err}")
    if _pending_drafts:
        st.warning(
            f"⚠️ **{len(_pending_drafts)} bozze NON inviate** per il periodo: queste modifiche NON verranno "
            "usate per i turni finché il medico non preme Salva.",
            icon="📝",
        )
        st.dataframe(
            pd.DataFrame([
                {
                    "Medico": p["doctor"],
                    "Mese": p["month"],
                    "Da aggiungere": p["added"],
                    "Da rimuovere": p["removed"],
                    "Ultima modifica": _local_hhmm(p["updated_at"]),
                }
                for p in _pending_drafts
            ]),
            hide_index=True,
            use_container_width=True,
        )
    _queued_saves = _save_queue().pending()
    if _queued_saves:
        st.error(
            f"⏳ **{len(_queued_saves)} salvataggi IN CODA** (GitHub saturo): non sono ancora registrati e "
            "NON verranno usati se generi adesso. Attendi qualche minuto che la coda si svuoti.",
            icon="⏳",
        )
        st.dataframe(
            pd.DataFrame([
                {
                    "Medico": j.doctor,
                    "Mesi": ", ".join(f"{int(yy):04d}-{int(mm):02d}" for (yy, mm) in sorted(j.entries_by_month)),
                    "Inviato alle": datetime.fromtimestamp(j.created_at, receipts.LOCAL_TZ).strftime("%d/%m %H:%M"),
                    "Tentativi": j.attempts,
                    "Ultimo errore": j.last_error,
                }
                for j in _queued_saves
            ]),
            hide_index=True,
            use_container_width=True,
        )
    if st.button("🔄 Ricontrolla bozze non inviate", key="refresh_pending_drafts"):
        load_pending_drafts_summary.clear()
        _rerun_fragment_or_app()

    # Step 1b-7: opzioni generazione applicate in blocco
    _options_key = f"admin_generation_options_{mk}"
    _options_state = st.session_state.get(_options_key)
    if not isinstance(_options_state, dict):
        _options_state = {}

    _disabled_excluded_cols = {"AD", "AE", "AF", "AG"}
    _disabled_col_labels = {
        f"{_col} · {_name}": _col
        for _col, _name in sorted((cfg_admin.get("columns") or {}).items())
        if _col not in _disabled_excluded_cols
    }
    _disabled_options = list(_disabled_col_labels.keys())
    _disabled_applied = [
        label for label in (_options_state.get("disabled_labels") or [])
        if label in _disabled_col_labels
    ]

    # Auto-carryover da storico: chi ha fatto notte l’ultimo giorno del mese precedente
    _carry_default = []
    try:
        _prev_period_day = period_start - timedelta(days=1)
        _mk_minus1 = f"{_prev_period_day.year:04d}-{_prev_period_day.month:02d}"
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

    _fixed_rows_default = []
    for _r in _options_state.get("fixed_assignments", []) or []:
        _dt = _period_date_or_default(_r.get("date"), period_start)
        _col = str(_r.get("column") or "J")
        _label = next((lb for lb, c in SHIFT_LABELS_ADMIN.items() if c == _col), "Notte (J)")
        _fixed_rows_default.append({
            "Medico": _r.get("doctor", doctors[0] if doctors else ""),
            "Giorno": _dt,
            "Turno/Colonna": _label,
        })
    _fixed_df_default = pd.DataFrame(_fixed_rows_default, columns=["Medico", "Giorno", "Turno/Colonna"])

    _V_WEEKDAYS = {0: "Lunedì", 2: "Mercoledì", 4: "Venerdì"}
    _weeks_v: dict = {}
    for _dd in _period_dates:
        _wd = _dd.weekday()
        if _wd in _V_WEEKDAYS:
            _iso_w = _dd.isocalendar()[:2]
            _weeks_v.setdefault(_iso_w, {})[_wd] = _dd

    _DOW_NAMES = {0: "Lunedì", 1: "Martedì", 2: "Mercoledì", 3: "Giovedì", 4: "Venerdì", 5: "Sabato", 6: "Domenica"}
    _weeks_j: dict = {}
    for _dd in _period_dates:
        _iso_w = _dd.isocalendar()[:2]
        _weeks_j.setdefault(_iso_w, {})[_dd.weekday()] = _dd

    _generation_memory_loaded = gmem.empty_memory()
    _generation_memory_sha = None
    _generation_memory_load_error = None
    try:
        _generation_memory_loaded, _generation_memory_sha = load_generation_memory_from_github_st()
    except Exception as _gme:
        _generation_memory_load_error = _gme
    _generation_versions = list((_generation_memory_loaded or {}).get("versions") or [])
    _version_rows = []
    for _v in _generation_versions:
        _assignments = _v.get("assignments") if isinstance(_v.get("assignments"), dict) else {}
        _n_days = len(_assignments)
        _n_slots = 0
        for _by_col in _assignments.values():
            if isinstance(_by_col, dict):
                _n_slots += sum(len(_docs) for _docs in _by_col.values() if isinstance(_docs, list))
        _version_rows.append({
            "Attiva": bool(_v.get("active", True)),
            "Etichetta": _v.get("label", ""),
            "Dal": _v.get("start_date", ""),
            "Al": _v.get("end_date", ""),
            "Creata": _v.get("created_at", ""),
            "Giorni": _n_days,
            "Slot": _n_slots,
            "ID": _v.get("id", ""),
        })
    _version_label_by_id = {
        str(row["ID"]): f"{row['Etichetta']} · {row['Dal']}–{row['Al']} · {str(row['ID'])[:8]}"
        for row in _version_rows
        if str(row.get("ID") or "")
    }
    _default_selected_ids = [
        str(row["ID"])
        for row in _version_rows
        if bool(row.get("Attiva", True)) and str(row.get("ID") or "")
    ]

    with st.form(f"admin_generation_options_form_{mk}"):
        st.markdown("### Opzioni generazione")
        st.caption("Fai tutte le selezioni qui dentro. Streamlit invia tutto in un solo run quando premi Applica opzioni o Genera turni.")

        with st.expander("1) Colonne da lasciare vuote nel periodo", expanded=False):
            st.info(
                "Seleziona le colonne da sospendere solo per questa generazione. "
                "Il solver non creerà quei turni e le relative celle resteranno vuote nell'Excel.",
                icon="🧯",
            )
            _disabled_selected_labels_draft = st.multiselect(
                "Colonne sospese",
                options=_disabled_options,
                default=_disabled_applied,
                key=f"disabled_columns_draft_{mk}",
                help="Esempio agosto: puoi sospendere ambulatori/servizi ridotti per ferie. Non modifica YAML o pool salvato.",
            )

        st.markdown("### 2) Indisponibilità")
        _unav_options = ["Nessuna", "Carica file manuale", "Usa archivio (privacy)"]
        _unav_default = str(_options_state.get("unav_mode") or "Usa archivio (privacy)")
        if _unav_default not in _unav_options:
            _unav_default = "Usa archivio (privacy)"
        unav_mode_draft = st.radio(
            "Fonte indisponibilità",
            _unav_options,
            index=_unav_options.index(_unav_default),
            horizontal=True,
            key=f"admin_unav_mode_draft_{mk}",
            help="Puoi caricare un file manuale, oppure usare l’archivio compilato dai medici.",
        )
        unav_upload_draft = st.file_uploader(
            "Carica indisponibilità manuale (usata solo se selezioni Carica file manuale)",
            type=["xlsx", "csv", "tsv"],
            key=f"unav_upload_draft_{mk}",
        )

        st.markdown("### 3) Vincolo post-notte a cavallo mese")
        _manual_default = [d for d in (_options_state.get("manual_block") or _carry_default) if d in doctors]
        manual_block_draft = st.multiselect(
            "Medico/i da bloccare il Giorno 1",
            doctors,
            default=_manual_default,
            key=f"manual_block_draft_{mk}",
            help="Chi ha fatto NOTTE l’ultimo giorno del mese precedente. Pre-compilato dallo storico se disponibile.",
        )

        st.markdown("### 4) Assegnazioni fisse (opzionale)")
        fixed_df_draft = st.data_editor(
            _fixed_df_default,
            column_config={
                "Medico": st.column_config.SelectboxColumn("Medico", options=doctors, required=False),
                "Giorno": st.column_config.DateColumn("Giorno", min_value=period_start, max_value=period_end, format="DD/MM/YYYY"),
                "Turno/Colonna": st.column_config.SelectboxColumn("Turno/Colonna", options=list(SHIFT_LABELS_ADMIN.keys()), required=False),
            },
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            key=f"fixed_assignments_draft_{mk}",
        )

        st.markdown("### 5) Turno doppio Sala PM (V) — eccezioni settimanali")
        st.info(
            "Di default ogni venerdì ha il turno doppio in V. Puoi spostarlo su lunedì/mercoledì o scegliere nessun doppio.",
            icon="🔬",
        )
        _v_double_draft = {}
        _v_saved = _options_state.get("v_double") if isinstance(_options_state.get("v_double"), dict) else {}
        for _iso_w in sorted(_weeks_v.keys()):
            _wdays = _weeks_v[_iso_w]
            _fri = _wdays.get(4)
            _alt_days = {_wd: _dt for _wd, _dt in _wdays.items() if _wd != 4}
            if not _alt_days:
                continue
            _label_default = f"Venerdì {_fri.strftime('%d/%m') if _fri else '(fuori mese)'} — doppio (default)"
            _label_no_double = "Nessun doppio questa settimana — tutti singoli"
            _options_v = {_label_default: None}
            for _wd_alt, _dt_alt in sorted(_alt_days.items()):
                _options_v[f"{_V_WEEKDAYS[_wd_alt]} {_dt_alt.strftime('%d/%m')} — doppio (invece del venerdì)"] = str(_dt_alt)
            _options_v[_label_no_double] = "NO_DOUBLE"
            _prev = _v_saved.get(str(_iso_w))
            _prev_label = next((lb for lb, val in _options_v.items() if val == _prev), _label_default)
            _sel_label = st.selectbox(
                f"Settimana {_iso_w[1]} ({min(_wdays.values()).strftime('%d/%m')}–{max(_wdays.values()).strftime('%d/%m')})",
                list(_options_v.keys()),
                index=list(_options_v.keys()).index(_prev_label),
                key=f"v_double_draft_{mk}_{_iso_w[0]}_{_iso_w[1]}",
            )
            _v_double_draft[str(_iso_w)] = _options_v[_sel_label]

        _rJ_admin = (cfg_admin.get("rules") or {}).get("J") or {}
        _j_blank_draft = {}
        _j_non_default_draft = {}
        if _rJ_admin.get("thursday_blank"):
            st.markdown("### 6) Giorni vuoti Notte (J) — eccezioni settimanali")
            st.info(
                "Di default ogni giovedì la colonna J è vuota. Qui puoi scegliere nessuno, uno o più giorni vuoti per settimana.",
                icon="🌙",
            )
            _j_saved = _options_state.get("j_blank") if isinstance(_options_state.get("j_blank"), dict) else {}
            for _iso_w in sorted(_weeks_j.keys()):
                _wdays = _weeks_j[_iso_w]
                _thu = _wdays.get(3)
                if _thu is None and not any(d in _wdays for d in range(0, 3)):
                    continue
                _opt_labels = {}
                for _wd in sorted(_wdays.keys()):
                    _dt_opt = _wdays[_wd]
                    _suffix = " (default)" if _wd == 3 else ""
                    _opt_labels[f"{_DOW_NAMES[_wd]} {_dt_opt.strftime('%d/%m')}{_suffix}"] = str(_dt_opt)
                _default_vals = [str(_thu)] if _thu is not None else []
                _prev_vals = _j_saved.get(str(_iso_w), _default_vals)
                _label_by_val = {v: k for k, v in _opt_labels.items()}
                _prev_labels = [_label_by_val[v] for v in _prev_vals if v in _label_by_val]
                _week_label = f"Settimana {_iso_w[1]} ({min(_wdays.values()).strftime('%d/%m')}–{max(_wdays.values()).strftime('%d/%m')}) — notti J vuote"
                _sel_labels = st.multiselect(
                    _week_label,
                    list(_opt_labels.keys()),
                    default=_prev_labels,
                    key=f"j_blank_draft_{mk}_{_iso_w[0]}_{_iso_w[1]}",
                )
                _j_blank_draft[str(_iso_w)] = sorted(_opt_labels[lb] for lb in _sel_labels)
                if sorted(_j_blank_draft[str(_iso_w)]) != sorted(_default_vals):
                    _j_non_default_draft[str(_iso_w)] = sorted(_j_blank_draft[str(_iso_w)])

            if _j_non_default_draft:
                _j_non_default_summary = []
                for _iso_key, _vals in sorted(_j_non_default_draft.items()):
                    try:
                        _yr_s, _wk_s = str(_iso_key).strip("()").replace("'", "").split(",")[:2]
                        _wk_label = f"{int(_yr_s)}-W{int(_wk_s):02d}"
                    except Exception:
                        _wk_label = str(_iso_key)
                    _j_non_default_summary.append(f"{_wk_label}: {', '.join(_vals) if _vals else 'nessuna J vuota'}")
                st.warning(
                    "Stai applicando eccezioni J non-default: " + "; ".join(_j_non_default_summary),
                    icon="🌙",
                )
            _confirm_j_non_default_draft = st.checkbox(
                "Confermo le eccezioni J non-default elencate sopra",
                value=bool(_options_state.get("confirm_j_non_default", False)) and bool(_j_non_default_draft),
                key=f"confirm_j_non_default_draft_{mk}",
                disabled=not bool(_j_non_default_draft),
            )
        else:
            _confirm_j_non_default_draft = False

        st.markdown("### 7) Memoria generazioni")
        _use_generation_memory_draft = st.checkbox(
            "Usa generazioni precedenti attive per ricalibrare quote/spaziature",
            value=bool(_options_state.get("use_generation_memory", True)),
            key=f"use_generation_memory_draft_{mk}",
            help="Per periodi parziali dello stesso mese, il solver conta i turni già generati prima del periodo corrente.",
        )
        _store_generation_memory_draft = st.checkbox(
            "Salva questa generazione in memoria dopo la creazione",
            value=bool(_options_state.get("store_generation_memory", False)),
            key=f"store_generation_memory_draft_{mk}",
            help="Crea una nuova versione. Le versioni precedenti restano disponibili e puoi attivarle/disattivarle.",
        )
        _generation_memory_label_draft = st.text_input(
            "Etichetta nuova versione",
            value=str(_options_state.get("generation_memory_label") or f"{period_start:%d/%m/%Y}-{period_end:%d/%m/%Y}"),
            key=f"generation_memory_label_draft_{mk}",
        )
        _selected_default_ids = _options_state.get("selected_generation_version_ids")
        if not isinstance(_selected_default_ids, list):
            _selected_default_ids = _default_selected_ids
        _selected_generation_version_ids_draft = st.multiselect(
            "Versioni da usare per questa generazione",
            options=list(_version_label_by_id.keys()),
            default=[vid for vid in _selected_default_ids if vid in _version_label_by_id],
            format_func=lambda vid: _version_label_by_id.get(str(vid), str(vid)),
            key=f"generation_memory_selected_versions_draft_{mk}",
            help="Le versioni non selezionate restano salvate ma non vengono usate.",
            disabled=not _use_generation_memory_draft,
        )

        _form_cols = st.columns([1, 1, 1, 2])
        with _form_cols[0]:
            apply_generation_options = st.form_submit_button("Applica opzioni")
        with _form_cols[1]:
            generate = st.form_submit_button("🚀 Genera turni", type="primary")
        with _form_cols[2]:
            reset_j_defaults = st.form_submit_button("Reset J default")

    if apply_generation_options or generate:
        _fixed_records = []
        if isinstance(fixed_df_draft, pd.DataFrame):
            for _, _row in fixed_df_draft.iterrows():
                _doc = str(_row.get("Medico") or "").strip()
                _day = _row.get("Giorno")
                _label = str(_row.get("Turno/Colonna") or "").strip()
                if not _doc or not _label:
                    continue
                if isinstance(_day, datetime):
                    _day = _day.date()
                if not isinstance(_day, date):
                    continue
                _fixed_records.append({"doctor": _doc, "date": str(_day), "column": SHIFT_LABELS_ADMIN.get(_label, "J")})
        _options_state = {
            "disabled_labels": list(_disabled_selected_labels_draft or []),
            "unav_mode": unav_mode_draft,
            "manual_block": list(manual_block_draft or []),
            "fixed_assignments": _fixed_records,
            "v_double": dict(_v_double_draft or {}),
            "j_blank": dict(_j_blank_draft or {}),
            "confirm_j_non_default": bool(_confirm_j_non_default_draft),
            "use_generation_memory": bool(_use_generation_memory_draft),
            "store_generation_memory": bool(_store_generation_memory_draft),
            "generation_memory_label": _generation_memory_label_draft,
            "selected_generation_version_ids": list(_selected_generation_version_ids_draft or []),
        }
        st.session_state[_options_key] = _options_state
        if generate and _j_non_default_draft and not _confirm_j_non_default_draft:
            st.error("Generazione bloccata: conferma le eccezioni J non-default oppure usa Reset J default.")
            generate = False
    elif reset_j_defaults:
        _options_state = dict(_options_state or {})
        _options_state["j_blank"] = {}
        _options_state["confirm_j_non_default"] = False
        st.session_state[_options_key] = _options_state
        for _iso_w in sorted(_weeks_j.keys()):
            st.session_state.pop(f"j_blank_draft_{mk}_{_iso_w[0]}_{_iso_w[1]}", None)
        st.session_state.pop(f"confirm_j_non_default_draft_{mk}", None)
        st.success("J ripristinata al default: solo giovedì vuoto per ogni settimana.")
        _rerun_fragment_or_app()
    else:
        generate = False

    disabled_columns_list = sorted(
        _disabled_col_labels[label]
        for label in (_options_state.get("disabled_labels") or [])
        if label in _disabled_col_labels
    )
    if disabled_columns_list:
        _critical_disabled = [c for c in disabled_columns_list if c in {"C", "D", "E", "H", "I", "J"}]
        if _critical_disabled:
            st.warning(
                "Colonne di guardia/continuità sospese: " + ", ".join(_critical_disabled),
                icon="⚠️",
            )
        st.caption("Colonne sospese applicate: " + ", ".join(disabled_columns_list))

    unav_mode = str(_options_state.get("unav_mode") or "Usa archivio (privacy)")
    unav_upload = unav_upload_draft if unav_mode == "Carica file manuale" else None
    use_archive = (unav_mode == "Usa archivio (privacy)")

    manual_block = [d for d in (_options_state.get("manual_block") or []) if d in doctors]
    carryover_by_month = {}
    _carryover_month_key = f"{period_start.year:04d}-{period_start.month:02d}"
    if manual_block:
        carryover_by_month[_carryover_month_key] = {"blocked_day1_doctors": list(dict.fromkeys(manual_block))}

    fixed_assignments_list = [
        {"doctor": r["doctor"], "date": r["date"], "column": r["column"]}
        for r in (_options_state.get("fixed_assignments") or [])
        if r.get("doctor") and r.get("date") and r.get("column")
    ]
    if fixed_assignments_list:
        st.caption(f"Assegnazioni fisse applicate: {len(fixed_assignments_list)}")

    _v_double_overrides_list = []
    for _iso_key, _val in sorted(((_options_state.get("v_double") or {}) if isinstance(_options_state.get("v_double"), dict) else {}).items()):
        if _val == "NO_DOUBLE":
            try:
                _yr_s, _wk_s = str(_iso_key).strip("()").replace("'", "").split(",")[:2]
                _v_double_overrides_list.append(f"NODOUBLE:{int(_yr_s)}:{int(_wk_s)}")
            except Exception:
                pass
        elif _val:
            _v_double_overrides_list.append(str(_val))
    if _v_double_overrides_list:
        st.caption("Eccezioni V applicate: " + ", ".join(_v_double_overrides_list))

    _j_blank_week_overrides = {}
    _rJ_admin = (cfg_admin.get("rules") or {}).get("J") or {}
    if _rJ_admin.get("thursday_blank"):
        for _iso_key, _vals in sorted(((_options_state.get("j_blank") or {}) if isinstance(_options_state.get("j_blank"), dict) else {}).items()):
            try:
                _yr_s, _wk_s = str(_iso_key).strip("()").replace("'", "").split(",")[:2]
                _wk_str = f"{int(_yr_s)}-W{int(_wk_s):02d}"
            except Exception:
                continue
            _vals = sorted(str(v) for v in (_vals or []))
            _thu_default = []
            for _iso_w, _wdays in _weeks_j.items():
                if str(_iso_w) == str(_iso_key) and _wdays.get(3) is not None:
                    _thu_default = [str(_wdays[3])]
                    break
            if _vals != sorted(_thu_default):
                _j_blank_week_overrides[_wk_str] = _vals
        if _j_blank_week_overrides:
            st.warning(
                "Eccezioni J non-default applicate: "
                + ", ".join(f"{k}: {', '.join(v) if v else 'nessun vuoto'}" for k, v in _j_blank_week_overrides.items()),
                icon="🌙",
            )

    _use_generation_memory = bool(_options_state.get("use_generation_memory", True))
    _store_generation_memory = bool(_options_state.get("store_generation_memory", False))
    _generation_memory_label = str(_options_state.get("generation_memory_label") or f"{period_start:%d/%m/%Y}-{period_end:%d/%m/%Y}")
    _selected_generation_version_ids = gmem.resolve_selected_version_ids(
        _options_state.get("selected_generation_version_ids"),
        _default_selected_ids,
    )
    if not _use_generation_memory:
        _selected_generation_version_ids = set()
    # Festivi locali (es. 3 giugno Messina) contano come festivi anche nella memoria.
    try:
        _memory_festive_dates = list(cfg_admin.get("festivi_extra") or []) + list(
            tg.load_turni_festivi().get("festivi_extra") or []
        )
    except Exception:
        _memory_festive_dates = list(cfg_admin.get("festivi_extra") or [])

    with st.expander("📜 Log inserimenti/modifiche indisponibilità (Audit)", expanded=False):
        st.caption("Il log resta separato dalle opzioni di generazione per evitare altri submit del pannello principale.")
        mk_log = f"{int(year)}-{int(month):02d}"
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
                st.dataframe(df_audit.head(200), use_container_width=True, hide_index=True)
            except Exception:
                pass

    with st.expander("Versioni salvate", expanded=False):
        if _generation_memory_load_error:
            st.warning(f"Memoria generazioni non leggibile: {_generation_memory_load_error}")
        elif not _version_rows:
            st.caption("Nessuna generazione salvata.")
        else:
            _versions_df = pd.DataFrame(_version_rows)
            with st.form(f"generation_memory_versions_form_{mk}"):
                _edited_versions = st.data_editor(
                    _versions_df,
                    column_config={
                        "Attiva": st.column_config.CheckboxColumn("Attiva"),
                        "Etichetta": st.column_config.TextColumn("Etichetta", disabled=True),
                        "Dal": st.column_config.TextColumn("Dal", disabled=True),
                        "Al": st.column_config.TextColumn("Al", disabled=True),
                        "Creata": st.column_config.TextColumn("Creata", disabled=True),
                        "Giorni": st.column_config.NumberColumn("Giorni", disabled=True),
                        "Slot": st.column_config.NumberColumn("Slot", disabled=True),
                        "ID": st.column_config.TextColumn("ID", disabled=True),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"generation_memory_versions_{mk}",
                    num_rows="fixed",
                )
                _save_versions = st.form_submit_button("💾 Salva stato versioni")
            if _save_versions:
                try:
                    _active_by_id = {
                        str(row.get("ID") or ""): bool(row.get("Attiva", True))
                        for _, row in _edited_versions.iterrows()
                    }
                    _fresh_mem, _fresh_sha = load_generation_memory_from_github_st()
                    _to_save = gmem.set_version_active(_fresh_mem, _active_by_id)
                    save_generation_memory_to_github_st(_to_save, _fresh_sha)
                    st.success("Stato versioni salvato.")
                    _rerun_fragment_or_app()
                except Exception as _e:
                    st.error(f"Errore salvataggio memoria generazioni: {_e}")

            try:
                _hist_preview, _ = _load_shift_history()
                _finalized_preview = set((_hist_preview or {}).keys())
            except Exception:
                _finalized_preview = set()
            _prior_preview = gmem.build_solver_prior_usage(
                _generation_memory_loaded,
                period_start,
                period_end,
                finalized_months=_finalized_preview,
                selected_version_ids=_selected_generation_version_ids,
                festive_dates=_memory_festive_dates,
            )
            _used = [v for v in _prior_preview.get("versions_used", []) if v]
            if _used:
                st.caption(f"Per questo periodo verrebbero usate {len(_used)} versioni precedenti selezionate.")
            else:
                st.caption("Per questo periodo non risultano versioni precedenti da conteggiare.")

    st.divider()

    if generate:
        t0 = time.time()
        status = st.status("Preparazione…", expanded=True)
        try:
            with tempfile.TemporaryDirectory() as td:
                td = Path(td)

                status.update(label="Preparazione template…", state="running")
                template_path = td / f"turni_{mk}.xlsx"
                if period_mode == "Mese intero":
                    tg.create_month_template_xlsx(
                        rules_path,
                        int(year),
                        int(month),
                        out_path=template_path,
                    )
                else:
                    tg.create_period_template_xlsx(
                        rules_path,
                        period_start,
                        period_end,
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
                    def _rows_in_period(_rows):
                        _out = []
                        for _r in _rows:
                            try:
                                _d = ustore.parse_iso_date(_r.get("date", ""))
                            except Exception:
                                continue
                            if period_start <= _d <= period_end:
                                _out.append(_r)
                        return _out

                    store_rows_1, sha1 = load_store_from_github()
                    rows_month = _rows_in_period(store_rows_1)
                    unav_path = td / "unavailability_from_store.xlsx"
                    xlsx_utils.build_unavailability_xlsx(rows_month, DEFAULT_UNAV_TEMPLATE, unav_path)

                    # Double-check solo in modalità legacy (SHA disponibile)
                    if sha1 is not None:
                        store_rows_2, sha2 = load_store_from_github()
                        if sha1 and sha2 and sha2 != sha1:
                            rows_month = _rows_in_period(store_rows_2)
                            xlsx_utils.build_unavailability_xlsx(rows_month, DEFAULT_UNAV_TEMPLATE, unav_path)
                            st.caption("Archivio indisponibilità aggiornato durante la preparazione: ricaricata l’ultima versione.")

                    st.caption(f"Archivio indisponibilità: {len(rows_month)} righe per {period_start:%d/%m/%Y}–{period_end:%d/%m/%Y}")

                status.update(label="Generazione turni…", state="running")
                out_path = td / f"output_{mk}.xlsx"

                # Carica preferenze di disponibilità da GitHub
                try:
                    _avail_all, _ = load_avail_store_from_github()
                    _avail_month = []
                    for _r in _avail_all:
                        try:
                            _d = ustore.parse_iso_date(_r.get("date", ""))
                        except Exception:
                            continue
                        if period_start <= _d <= period_end:
                            _avail_month.append(_r)
                    all_avail_prefs = [
                        {"doctor": r["doctor"], "date": r["date"], "shift": r["shift"],
                         "priority": r.get("priority", "media")}
                        for r in _avail_month
                    ]
                except Exception as _e:
                    all_avail_prefs = []
                    st.warning(f"Impossibile caricare preferenze da GitHub: {_e}")

                # Pool config overlay (se presente su GitHub)
                _pool_cfg_for_solver, _ = load_pool_config_from_github_st()

                _hist_data2, _ = _load_shift_history()
                _finalized_months_for_memory = set((_hist_data2 or {}).keys())

                _prior_usage_for_solver = None
                _gen_mem_fresh = None
                if _use_generation_memory:
                    status.update(label="Carico memoria generazioni…", state="running")
                    _gen_mem_fresh, _gen_mem_sha_fresh = load_generation_memory_from_github_st()
                    _prior_usage_for_solver = gmem.build_solver_prior_usage(
                        _gen_mem_fresh,
                        period_start,
                        period_end,
                        finalized_months=_finalized_months_for_memory,
                        selected_version_ids=_selected_generation_version_ids,
                        festive_dates=_memory_festive_dates,
                    )
                    _used_versions = [v for v in _prior_usage_for_solver.get("versions_used", []) if v]
                    if _used_versions:
                        st.caption(f"Memoria generazioni: uso {len(_used_versions)} versione/i precedenti selezionate.")

                # Storico aggregato effettivo per il solver:
                # definitivo importato + generazioni attive non ancora sostituite da definitivo.
                _effective_hist_data = dict(_hist_data2 or {})
                _gen_hist_for_solver = {}
                if _use_generation_memory and _gen_mem_fresh is not None:
                    _gen_hist_for_solver = gmem.memory_to_shift_history(
                        _gen_mem_fresh,
                        start_date=period_start,
                        end_date=period_end,
                        finalized_months=set(_effective_hist_data.keys()),
                        valid_doctors=set(doctors) if doctors else None,
                        selected_version_ids=_selected_generation_version_ids,
                    )
                    for _hmk, _hstats in _gen_hist_for_solver.items():
                        if _hmk not in _effective_hist_data:
                            _effective_hist_data[_hmk] = _hstats
                _hist_agg_for_solver = sh.aggregate_multi_month(_effective_hist_data) if _effective_hist_data else None
                if _gen_hist_for_solver:
                    st.caption(
                        "Memoria storica effettiva: include generazioni attive per "
                        + ", ".join(sorted(_gen_hist_for_solver.keys()))
                    )

                status.update(label="Generazione turni…", state="running")
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
                    disabled_columns=disabled_columns_list if disabled_columns_list else None,
                    historical_stats=_hist_agg_for_solver,
                    pool_config=_pool_cfg_for_solver if _pool_cfg_for_solver else None,
                    prior_usage=_prior_usage_for_solver,
                )

                _saved_generation_memory = None
                if _store_generation_memory:
                    status.update(label="Salvo memoria generazione…", state="running")
                    _assignments_for_memory = gmem.parse_generated_xlsx_assignments(out_path)
                    if _assignments_for_memory:
                        _mem_latest, _mem_latest_sha = load_generation_memory_from_github_st()
                        _version_id = str(uuid.uuid4())
                        _mem_to_save = gmem.append_version(
                            _mem_latest,
                            version_id=_version_id,
                            label=_generation_memory_label or f"{period_start:%d/%m/%Y}-{period_end:%d/%m/%Y}",
                            start_date=period_start,
                            end_date=period_end,
                            assignments=_assignments_for_memory,
                            active=True,
                        )
                        save_generation_memory_to_github_st(_mem_to_save, _mem_latest_sha)
                        _saved_generation_memory = {
                            "version_id": _version_id,
                            "days": len(_assignments_for_memory),
                        }
                    else:
                        _saved_generation_memory = {
                            "version_id": None,
                            "days": 0,
                            "warning": "Nessuna assegnazione letta dall'Excel generato.",
                        }

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
                    "generation_memory": {
                        "used": bool(_use_generation_memory),
                        "versions_used": (_prior_usage_for_solver or {}).get("versions_used", []) if _prior_usage_for_solver else [],
                        "saved": _saved_generation_memory,
                    },
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
        _gm_last = last.get("generation_memory") if isinstance(last.get("generation_memory"), dict) else {}
        if _gm_last:
            _gm_parts = []
            if _gm_last.get("used"):
                _gm_parts.append(f"memoria usata: {len(_gm_last.get('versions_used') or [])} versioni")
            _gm_saved = _gm_last.get("saved") if isinstance(_gm_last.get("saved"), dict) else None
            if _gm_saved and _gm_saved.get("version_id"):
                _gm_parts.append(f"salvata nuova versione ({_gm_saved.get('days', 0)} giorni)")
            elif _gm_saved and _gm_saved.get("warning"):
                st.warning(str(_gm_saved.get("warning")))
            if _gm_parts:
                st.caption("Memoria generazioni: " + " · ".join(_gm_parts))
        # Mostra solver_error se INFEASIBLE (per diagnostica)
        if str(_stats.get("status","")).upper() == "INFEASIBLE":
            _month_stats = (_stats.get("months") or {}).get(mk, {}) or {}
            _serr = _month_stats.get("solver_error") or _stats.get("solver_error") or "(nessun dettaglio)"
            st.error(f"**Errore solver:** {_serr}")

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
            blocked = (carryover_by_month.get(_carryover_month_key, {}) or {}).get("blocked_day1_doctors", [])
            st.markdown(
                f'<div class="kpi"><b>Carryover</b><br>{len(blocked)} bloccati Giorno 1</div>',
                unsafe_allow_html=True,
            )

        # Avvisi quota rilassata (J, festivi, pool_quota_overrides)
        if isinstance(_stats, dict) and _stats.get("warnings"):
            st.warning(
                "⚠️ **Quote rilassate per indisponibilità** — il solver ha adattato "
                "alcune quote perché non erano raggiungibili:\n\n"
                + "\n".join(f"- {w}" for w in _stats["warnings"]),
            )
        if isinstance(_stats, dict) and _stats.get("C_reperibilita_diag"):
            _cdiag = _stats["C_reperibilita_diag"]
            _relax_warns = _cdiag.get("relaxation_warnings") if isinstance(_cdiag, dict) else None
            if _relax_warns:
                st.warning(
                    "⚠️ **Reperibilità (C) generata con vincoli rilassati** — "
                    "verifica il risultato:\n\n" + "\n".join(f"- {w}" for w in _relax_warns),
                )
            with st.expander("Dettagli Reperibilità (C)"):
                st.json(_cdiag)


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

# Aggiungi medici attivi dal pool_config non già presenti nel YAML
# (permette di aggiungere nuovi medici dalla GUI senza toccare Regole_Turni.yml)
try:
    _pc_login = _pool_config_for_doctor_list()
    if _pc_login:
        _existing_dn = {tg.norm_name(d) for d in doctors_default}
        _new_from_pc = [
            d for d, dc in _pc_login.get("doctors", {}).items()
            if dc.get("active", True)
            and tg.norm_name(d) not in _existing_dn
            and d != "Recupero"
        ]
        if _new_from_pc:
            doctors_default = sorted(
                doctors_default + _new_from_pc,
                key=lambda s: (s == "Recupero", s.lower()),
            )
except Exception:
    pass

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
            # le modifiche non inviate restano nella bozza sul server, non in memoria
            _reset_doctor_editor_state(old_doctor)
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
                    # Una bozza rimasta in memoria da una sessione precedente (es. dopo
                    # un kick-out) non deve riapparire sopra i dati aggiornati del server.
                    _reset_doctor_editor_state(doctor)
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

    # Completa un salvataggio eventualmente interrotto da un doppio tap
    # (audit, mail, bozza) prima di mostrare i dati.
    _process_unav_outbox(doctor, selected)
    _show_save_queue_status(doctor, selected)

    refresh_baseline = st.button(
        "🔄 Ricarica dati",
        help="Ricarica i dati dal server. Le modifiche non inviate restano nella bozza e potrai recuperarle.",
    )

    if refresh_baseline:
        # Le modifiche non inviate vanno prima in bozza, poi si riparte dal server.
        _draft_flush(doctor, force=True)
        clear_doctor_baseline()
        _reset_doctor_editor_state(doctor)
        st.rerun()

    try:
        baseline = get_or_load_doctor_baseline(doctor, selected, force_reload=bool(refresh_baseline))
        store_rows = list(baseline.get("rows") or [])
    except github_utils.GithubUnavailable as e:
        st.warning(
            f"⏳ GitHub è momentaneamente saturo e non riesco a leggere le tue indisponibilità ({e}). "
            "Riprova tra un minuto con 🔄 Ricarica dati: nessun dato è andato perso."
        )
        st.stop()
    except Exception as e:
        st.error(f"Errore accesso archivio indisponibilità: {e}")
        st.stop()
    if baseline.get("stale"):
        st.info(
            f"ℹ️ GitHub è momentaneamente saturo: vedi gli ultimi dati letti ({_local_hhmm(baseline.get('loaded_at'))}). "
            "Puoi modificare e salvare: il salvataggio ricontrolla il server o va in coda."
        )

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
            st.caption(
                "Inserisci i giorni in cui NON puoi lavorare. Le modifiche vengono salvate in bozza "
                "automaticamente, ma valgono per i turni solo dopo 💾 Salva indisponibilità."
            )
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


            # --- Editor righe (UI robusta: niente st.data_editor) ---
            rows_key = f"unav_rows_{doctor}_{yy}_{mm}"
            mk_cur = f"{yy:04d}-{mm:02d}"
            official_sig = ustore.month_signature(store_rows, doctor, yy, mm)
            base_sig_key = f"{rows_key}__base_sig"
            draft_offer_key = f"{rows_key}__draft_offer"

            def _editor_rows(items: list[dict]) -> list[dict]:
                return [
                    {
                        "id": str(uuid.uuid4()),
                        "Data": _r.get("Data"),
                        "Fascia": _r.get("Fascia") or "Mattina",
                        "Note": _r.get("Note", ""),
                    }
                    for _r in items
                ]

            # Initialize per-month rows state once (or after login/reload):
            # server data, or the unsent draft if the doctor left changes pending.
            if rows_key not in st.session_state:
                start = {"action": "official"}
                try:
                    start = udrafts.editor_start(_session_draft(doctor), mk_cur, official_sig)
                except Exception as _draft_err:
                    st.caption(f"Bozza non leggibile ({_draft_err}): mostro i dati registrati.")
                if start["action"] == "resume":
                    init = [{"Data": d, "Fascia": sh, "Note": note} for d, sh, note in start["entries"]]
                    set_unav_flash(
                        doctor,
                        "info",
                        f"Ho ripristinato le modifiche NON inviate salvate in bozza il {_local_hhmm(start['updated_at'])}. "
                        "Controllale e premi 💾 Salva indisponibilità per inviarle.",
                    )
                elif start["action"] == "offer":
                    st.session_state[draft_offer_key] = start
                st.session_state[rows_key] = _editor_rows(init)
                st.session_state[base_sig_key] = official_sig
                _draft_mark_written(
                    doctor,
                    mk_cur,
                    [(r["Data"], r["Fascia"], r.get("Note", "")) for r in init if isinstance(r.get("Data"), date)],
                )

            _offer = st.session_state.get(draft_offer_key)
            if _offer and unav_open:
                _offer_list = ", ".join(f"{d:%d/%m} {sh}" for d, sh, _n in _offer.get("entries") or []) or "nessuna riga"
                st.warning(
                    f"Hai una **bozza NON inviata** del {_local_hhmm(_offer.get('updated_at'))} "
                    f"({_offer_list}), ma nel frattempo le indisponibilità registrate sono cambiate "
                    "(ad es. da un altro dispositivo). Qui sotto vedi quelle registrate: vuoi recuperare la bozza?"
                )
                _oc1, _oc2 = st.columns(2)
                with _oc1:
                    if st.button("↩️ Recupera bozza", key=f"{rows_key}__draft_recover", use_container_width=True):
                        _draft_rows = [{"Data": d, "Fascia": sh, "Note": n} for d, sh, n in _offer["entries"]]
                        st.session_state[rows_key] = _editor_rows(_draft_rows)
                        st.session_state.pop(draft_offer_key, None)
                        _draft_mark_written(doctor, mk_cur, _offer["entries"])
                        st.rerun()
                with _oc2:
                    if st.button("🗑️ Scarta bozza", key=f"{rows_key}__draft_discard", use_container_width=True):
                        _draft_discard_month(doctor, mk_cur)
                        st.session_state.pop(draft_offer_key, None)
                        st.rerun()

            first_day = date(yy, mm, 1)
            if mm == 12:
                last_day = date(yy + 1, 1, 1) - timedelta(days=1)
            else:
                last_day = date(yy, mm + 1, 1) - timedelta(days=1)

            if unav_open:
                # Nessuna riga precompilata: una riga "1 del mese, Mattina" lasciata lì
                # veniva salvata come indisponibilità vera.
                rows = list(st.session_state.get(rows_key) or [])

                with st.expander("📅 Aggiungi periodo ferie", expanded=False):
                    st.caption(
                        "Seleziona un intervallo di date e premi **Aggiungi giorni** per inserire ogni giorno "
                        "del periodo come riga con fascia *Ferie* nella tabella sottostante "
                        "(poi ricordati di premere **💾 Salva indisponibilità**). "
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

                if rows:
                    h1, h2, h3 = st.columns([2, 2, 1])
                    h1.markdown("**Data**")
                    h2.markdown("**Fascia**")
                    h3.markdown("**Rimuovi**")
                else:
                    st.info(
                        "Nessuna indisponibilità per questo mese. Usa **➕ Aggiungi riga** "
                        "o **📅 Aggiungi periodo ferie**."
                    )

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

            # ── Bozza automatica + stato "non inviato" ─────────────────────────
            _dstate = _draft_state(doctor)
            if ustore.entries_signature(entries_norm) != official_sig:
                _udiff = ustore.compute_month_diff(existing, entries_norm)
                _dstate["unsent"][mk_cur] = {"added": _udiff["added_count"], "removed": _udiff["removed_count"]}
            else:
                _dstate["unsent"].pop(mk_cur, None)
            if unav_open and not st.session_state.get(draft_offer_key):
                _draft_note_change(
                    doctor, mk_cur, entries_norm, st.session_state.get(base_sig_key, official_sig), _doctor_session_id
                )
                _draft_flush(doctor)
            if unav_open:
                _draft_status_and_flush(doctor, _doctor_session_id, mk_cur)

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
                        # Nessuna riga precompilata: verrebbe salvata come preferenza vera.
                        st.session_state[avail_key] = []

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
                            ]
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
            if not touch_doctor_session_lease(doctor, _sid):
                _logout_doctor(
                    "Impossibile salvare: la sessione è stata sostituita da un accesso dello stesso utente da un altro dispositivo/browser."
                )
                st.stop()
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

        base_signatures = {
            (yy, mm): st.session_state.get(
                f"unav_rows_{doctor}_{yy}_{mm}__base_sig",
                ustore.month_signature(store_rows, doctor, yy, mm),
            )
            for (yy, mm) in (normalized_entries_by_month or {})
        }

        save_cfg = _service_config()
        save_job = usvc.make_job(
            doctor=doctor,
            doctor_email=str(get_doctor_contact(doctor).get("email") or ""),
            entries_by_month=dict(normalized_entries_by_month or {}),
            base_signatures=base_signatures,
            origin=_sid,
        )
        save_status = None
        try:
            with st.spinner("Salvataggio in corso… non chiudere la pagina"):
                queue = _save_queue()
                queue.cfg = save_cfg
                save_status, save_result = queue.save_or_enqueue(save_job, max_wait=25)
                # Nessuna chiamata st.* tra il salvataggio e l'outbox: un doppio tap
                # può interrompere l'esecuzione e il resto va completato comunque.
                if save_status == "saved":
                    st.session_state[f"unav_outbox::{doctor}"] = {
                        "result": save_result,
                        "cfg": save_cfg,
                        "selected": list(selected),
                    }
        except ustore.MonthConflictError as e:
            _draft_flush(doctor, force=True)
            set_unav_flash(
                doctor,
                "error",
                f"❌ Salvataggio BLOCCATO: {e} Per non cancellare quei dati non ho sovrascritto nulla. "
                "Le tue modifiche sono al sicuro nella bozza: premi 🔄 Ricarica dati, controlla le "
                "indisponibilità aggiornate e scegli se recuperare la bozza.",
            )
            st.rerun()
        except Exception as e:
            _draft_flush(doctor, force=True)
            set_unav_flash(
                doctor,
                "error",
                f"Errore durante il salvataggio ❌ — {type(e).__name__}: {e}. "
                "Le modifiche NON sono state inviate (restano nella bozza).",
                details=(
                    "Se il problema persiste: ricarica la pagina e riprova.\n\n"
                    "Se vedi 404: (1) token senza accesso alla repo privata, "
                    "(2) owner/repo/branch/path errati, "
                    "(3) token non autorizzato SSO (se repo in Organization)."
                ),
            )
            st.rerun()
        if save_status == "queued":
            _draft_flush(doctor, force=True)
            set_unav_flash(
                doctor,
                "warning",
                "⏳ GitHub è momentaneamente saturo: il salvataggio è **IN CODA** e verrà registrato "
                "automaticamente appena possibile (di solito entro pochi minuti), anche se chiudi la pagina. "
                "Riceverai la mail di conferma solo a registrazione avvenuta: senza mail non è registrato. "
                "Non serve premere di nuovo Salva.",
            )
            st.rerun()
        _process_unav_outbox(doctor, selected)
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
    try:
        _pc_admin, _ = load_pool_config_from_github_st()
        if _pc_admin:
            _dn_set = {tg.norm_name(d) for d in doctors}
            _new = [
                d for d, dc in _pc_admin.get("doctors", {}).items()
                if dc.get("active", True)
                and tg.norm_name(d) not in _dn_set
                and d != "Recupero"
            ]
            if _new:
                doctors = sorted(
                    doctors + _new,
                    key=lambda s: (s == "Recupero", s.lower()),
                )
    except Exception:
        pass
    rules_path = DEFAULT_RULES_PATH

    # ── Ramo Configurazione ───────────────────────────────────────────────
    if mode == "🔧 Admin — Configurazione":
        st.markdown("### 🔧 Configurazione sistema")

        # Flash messages da operazioni precedenti (es. salvataggio pool)
        if "_cfg_flash" in st.session_state:
            _fk, _fm = st.session_state.pop("_cfg_flash")
            if _fk == "success":
                st.success(_fm, icon="✅")
            else:
                st.error(_fm)

        with st.expander("⚙️ Impostazioni indisponibilità", expanded=False):
            try:
                app_settings, app_settings_sha = _load_app_settings_uncached()
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

            st.markdown("**Copie della mail di resoconto** *(inviata al medico a ogni salvataggio con modifiche)*")
            _cc_default = "\n".join(app_settings.get("receipt_cc_emails") or [])
            new_receipt_cc_text = st.text_area(
                "Indirizzi in copia (uno per riga o separati da virgola)",
                value=_cc_default,
                height=80,
                key="receipt_cc_emails_input",
            )
            new_receipt_cc = receipts.parse_email_list(new_receipt_cc_text)
            _bad_cc = [a for a in new_receipt_cc if not receipts.recipients(a, [])[0]]
            if _bad_cc:
                st.warning("Indirizzi non validi (verranno ignorati): " + ", ".join(_bad_cc))

            if st.button("Salva impostazioni indisponibilità", type="primary"):
                settings_to_save = {
                    "unavailability_open": bool(new_open),
                    "max_unavailability_per_shift": int(new_max),
                    "max_availability_per_shift": int(new_max_avail),
                    "max_weekend_days": int(new_max_weekend),
                    "doctor_caps": {doc: int(v) for doc, v in new_doctor_caps.items()},
                    "receipt_cc_emails": [a for a in new_receipt_cc if a not in _bad_cc],
                    "updated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    "updated_by": "admin",
                }
                try:
                    save_app_settings_to_github(
                        settings_to_save,
                        app_settings_sha,
                        message=f"Update settings: open={bool(new_open)} max_unav={int(new_max)} max_avail={int(new_max_avail)} max_weekend={int(new_max_weekend)}",
                    )
                    st.session_state["_cfg_flash"] = ("success", "Impostazioni salvate ✅")
                    st.rerun()
                except Exception as e:
                    st.session_state["_cfg_flash"] = ("error", f"Errore salvataggio impostazioni su GitHub: {e}")
                    st.rerun()

        with st.expander("🗂️ Gestione admin indisponibilità / preferenze", expanded=False):
            render_admin_doctor_data_editor(doctors, date.today().year, date.today().month)

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
                _LIBRE_COLS = {"AD", "AE", "AF", "AG"}
                # AA copia automaticamente K+T; AC è sempre Migliorato (fixed) — non esporre nella GUI
                _AUTO_COLS = {"AA", "AC"}
                _all_cols = sorted(k for k in cfg_admin.get("columns", {}).keys() if k not in _LIBRE_COLS and k not in _AUTO_COLS)
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
                            "No sabato diurno": bool(_dc.get("exclude_saturday_day", False)),
                            "Festivi diurni": bool(_dc.get("festivi_diurni", True)),
                            "Festivi notti": bool(_dc.get("festivi_notti", True)),
                            "Universitario": bool(_dc.get("university_doctor")),
                            "Email": str(_dc.get("email") or ""),
                        })
                    _med_df = _pd_pool.DataFrame(_med_rows)
                    _edited_med = st.data_editor(
                        _med_df,
                        column_config={
                            "Medico": st.column_config.TextColumn("Medico", help="Cognome esatto (maiuscola iniziale)"),
                            "Attivo": st.column_config.CheckboxColumn("Attivo"),
                            "Reperibilità": st.column_config.CheckboxColumn("Reperibilità C"),
                            "No sabato diurno": st.column_config.CheckboxColumn("No sabato mattina/pomeriggio"),
                            "Festivi diurni": st.column_config.CheckboxColumn("Festivi diurni"),
                            "Festivi notti": st.column_config.CheckboxColumn("Festivi notti"),
                            "Universitario": st.column_config.CheckboxColumn("Universitario"),
                            "Email": st.column_config.TextColumn("Email", help="Email per OTP/PIN (salvata in doctor_contacts.yml)"),
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
                                "exclude_saturday_day": False,
                                "excluded_from_reperibilita": False,
                                "university_doctor": None, "column_overrides": {},
                                "email": None,
                            }
                        _draft_doctors[_dn]["active"] = bool(_row["Attivo"])
                        _draft_doctors[_dn]["excluded_from_reperibilita"] = not bool(_row["Reperibilità"])
                        _draft_doctors[_dn]["exclude_saturday_day"] = bool(_row["No sabato diurno"])
                        _draft_doctors[_dn]["festivi_diurni"] = bool(_row["Festivi diurni"])
                        _draft_doctors[_dn]["festivi_notti"] = bool(_row["Festivi notti"])
                        _is_uni = bool(_row["Universitario"])
                        if _is_uni:
                            _draft_doctors[_dn]["university_doctor"] = {"ratio": 0.6}
                        else:
                            _draft_doctors[_dn]["university_doctor"] = None
                        _email_val = str(_row.get("Email") or "").strip()
                        _draft_doctors[_dn]["email"] = _email_val if _email_val else None
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
                        {
                            "Col 1": c["columns"][0],
                            "Col 2": c["columns"][1],
                            "Modalità": "fallback"
                            if tuple(sorted(str(_cc).strip().upper() for _cc in c.get("columns", []))) == ("K", "T")
                            and c.get("mode") == "always"
                            else c["mode"],
                        }
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
                            if tuple(sorted([str(_c1).strip().upper(), str(_c2).strip().upper()])) == ("K", "T"):
                                _mode = "fallback"
                            _new_combos.append({"columns": [str(_c1), str(_c2)], "same_day": True, "mode": str(_mode)})
                    _draft["service_combinations"] = _new_combos

                    st.divider()
                    st.markdown("**Servizi indispensabili**")
                    st.info(
                        "Le colonne marcate come **indispensabili** hanno un fallback automatico: "
                        "se il pool primario è esaurito (per ferie, indisponibilità, smonti notte), "
                        "il solver può assegnare **qualsiasi medico disponibile** in quel turno. "
                        "Il medico di emergenza riceve una penalità alta — viene usato solo se "
                        "non esiste alternativa — ma la colonna non rimane vuota.",
                        icon="🚨",
                    )
                    _critical: dict = _draft.setdefault("critical_services", {})
                    # Colonne attualmente marcate come indispensabili (fallback=any)
                    _curr_critical_cols = sorted(
                        col for col, spec in _critical.items() if spec.get("fallback") == "any"
                    )
                    _selected_critical = st.multiselect(
                        "Colonne indispensabili (fallback → qualsiasi medico)",
                        options=_all_cols,
                        default=[c for c in _curr_critical_cols if c in _all_cols],
                        key="pool_critical_multisel",
                        help="Seleziona le colonne che non devono mai rimanere vuote. "
                             "D, F, H, I, J sono tipicamente indispensabili.",
                    )
                    # Aggiorna critical_services mantenendo eventuali fallback custom (lista medici)
                    _new_crit: dict = {}
                    for _col in _selected_critical:
                        # Se aveva già un fallback custom (lista), mantienilo; altrimenti "any"
                        _existing = _critical.get(_col, {})
                        if isinstance(_existing.get("fallback"), list):
                            _new_crit[_col] = _existing
                        else:
                            _new_crit[_col] = {"fallback": "any"}
                    _draft["critical_services"] = _new_crit

                st.divider()
                _audit = _pcs_ui.audit_pool_config(_draft, cfg_admin)
                with st.expander("🧪 Controllo configurazione prima del salvataggio", expanded=bool(_audit["errors"])):
                    if _audit["errors"]:
                        st.error("Errori bloccanti:\n\n" + "\n".join(f"- {e}" for e in _audit["errors"]))
                    else:
                        st.success("Nessun errore bloccante nella configurazione.")
                    if _audit["warnings"]:
                        st.warning("Avvisi:\n\n" + "\n".join(f"- {w}" for w in _audit["warnings"]))
                    _pool_preview = [
                        {
                            "Colonna": _col,
                            "Medici attivi nel pool": ", ".join(_docs) if _docs else "—",
                            "N": len(_docs),
                        }
                        for _col, _docs in sorted((_audit.get("preview") or {}).get("column_pools", {}).items())
                    ]
                    if _pool_preview:
                        st.dataframe(_pd_pool.DataFrame(_pool_preview), use_container_width=True, hide_index=True)

                _col_save, _col_reset = st.columns([3, 1])
                with _col_reset:
                    if st.button("↩️ Reset draft", key="btn_pool_reset"):
                        del st.session_state["pool_cfg_draft"]
                        st.rerun()
                with _col_save:
                    if st.button("💾 Salva configurazione pool", type="primary", key="btn_pool_save"):
                        import copy as _copy_save
                        from datetime import datetime as _dt_save, timezone as _tz_save
                        _to_save = _pcs_ui.normalize_pool_config(_copy_save.deepcopy(_draft))
                        _to_save["updated_at"] = _dt_save.now(_tz_save.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                        _to_save["updated_by"] = "admin"
                        _errs = _pcs_ui.validate_pool_config(_to_save, cfg_admin)
                        if _errs:
                            st.error("Configurazione non valida:\n\n" + "\n".join(f"- {e}" for e in _errs))
                        else:
                            _ok_save, _msg_save = save_pool_config_with_retry(_to_save, _pool_cfg_sha)
                            if _ok_save:
                                # Sincronizza email medici con doctor_contacts.yml
                                _sync_ok, _sync_msg = sync_pool_contacts_to_github(_to_save)
                                if not _sync_ok:
                                    st.warning(f"Pool salvato, ma sync contatti fallita: {_sync_msg}")
                                del st.session_state["pool_cfg_draft"]
                                st.session_state["_cfg_flash"] = ("success", _msg_save)
                                st.rerun()
                            else:
                                st.session_state["_cfg_flash"] = ("error", _msg_save)
                                st.rerun()

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

            _gen_hist_data = {}
            _gen_hist_error = None
            try:
                _gen_mem_hist, _ = load_generation_memory_from_github_st()
                _gen_hist_data = gmem.memory_to_shift_history(
                    _gen_mem_hist,
                    finalized_months=set((_hist_data or {}).keys()),
                    valid_doctors=set(doctors) if doctors else None,
                )
            except Exception as _e:
                _gen_hist_error = _e
                _gen_hist_data = {}

            _effective_hist_data = dict(_hist_data or {})
            for _hmk, _hstats in (_gen_hist_data or {}).items():
                if _hmk not in _effective_hist_data:
                    _effective_hist_data[_hmk] = _hstats

            if _gen_hist_error:
                st.warning(f"Memoria generazioni non leggibile nella vista storica: {_gen_hist_error}")

            if _effective_hist_data:
                _sorted_months = sorted(_effective_hist_data.keys())
                _final_months = sorted((_hist_data or {}).keys())
                _gen_months = sorted((_gen_hist_data or {}).keys())
                if _final_months:
                    st.caption(f"Mesi definitivi importati: {', '.join(_final_months)}")
                if _gen_months:
                    st.caption(
                        "Mesi provvisori da generazioni attive: "
                        + ", ".join(_gen_months)
                        + " (ignorati automaticamente se esiste il definitivo dello stesso mese)"
                    )
                _agg_hist = sh.aggregate_multi_month(_effective_hist_data)

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
                        _ms_sel = _effective_hist_data[_sel_mese]
                        _src = "definitivo" if _sel_mese in (_hist_data or {}) else "generazione attiva"
                        st.caption(f"Origine mese: **{_src}**")
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
                                    for _doc2, _ds2 in _effective_hist_data[_ml2].items():
                                        if _doc2 == "_meta": continue
                                        _j2 = _ds2.get("J", {})
                                        _evo.append({"Mese": _ml2, "Medico": _doc2, "Notti": _j2.get("total", 0) if isinstance(_j2, dict) else 0})
                                _fig = px.line(pd.DataFrame(_evo), x="Mese", y="Notti", color="Medico", title="Evoluzione notti per medico", markers=True)
                                _fig.update_layout(height=400); st.plotly_chart(_fig, use_container_width=True)
                            else:
                                st.info("Servono almeno 2 mesi per il grafico di evoluzione.")

                if _final_months:
                    with st.expander("🗑️ Rimuovi mese definitivo dallo storico", expanded=False):
                        _month_del = st.selectbox("Seleziona mese definitivo da rimuovere", _final_months, key="hist_del")
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

    render_admin_generate_panel(cfg_admin, doctors, rules_path)
