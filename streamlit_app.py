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
import yaml

# Local modules
import github_utils
import unavailability_store as ustore
import xlsx_utils

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

# ---- Indisponibilità: fasce ammesse e normalizzazione (per compatibilità con valori "storici") ----
FASCIA_OPTIONS = ["Mattina", "Pomeriggio", "Notte", "Diurno", "Tutto il giorno"]

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

    # unknown
    return "Tutto il giorno", True, True
# ---------------- Page config & style ----------------
st.set_page_config(
    page_title="Turni UTIC – Autogeneratore",
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
    }

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
def load_store_from_github() -> tuple[list[dict], str | None]:
    g = _github_cfg()
    if not (g.get("token") and g.get("owner") and g.get("repo") and g.get("path")):
        raise RuntimeError("Archivio indisponibilità: secrets GitHub non configurati.")
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
    """Return {doctor: {email, phone}} from a YAML file in the same GitHub repo."""
    g = _github_cfg()
    path = g.get("contacts_path") or "data/doctor_contacts.yml"
    gf = github_utils.get_file(
        owner=g["owner"], repo=g["repo"], path=path, token=g["token"], branch=g.get("branch","main")
    )
    if gf is None:
        return {}
    try:
        data = yaml.safe_load(gf.text) or {}
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
    """Concurrency-safe save for the shared GitHub CSV.

    Strategy:
      - Apply the doctor's month replacements on top of the latest store version.
      - If a concurrent save happens (SHA conflict), reload and retry with backoff.
    Returns: (audit_todo, final_sha)
    """
    last_err: Exception | None = None

    # Deterministic order
    months = sorted(normalized_entries_by_month.items(), key=lambda kv: (kv[0][0], kv[0][1]))

    for attempt in range(max_retries):
        # Use the previously loaded store for the first attempt to save one roundtrip.
        if attempt == 0 and initial_rows is not None:
            store_rows = list(initial_rows)
            store_sha = initial_sha
        else:
            store_rows, store_sha = load_store_from_github()

        if expected_signatures:
            stale_mk = _detect_stale_doctor_month(store_rows, doctor, expected_signatures)
            if stale_mk:
                raise RuntimeError(
                    "Conflitto di aggiornamento: le tue indisponibilità "
                    f"per {stale_mk} sono state modificate dopo il caricamento. "
                    "Ricarica la pagina e riprova."
                )

        new_rows = list(store_rows)
        audit_todo: list[tuple[str, dict]] = []

        for (yy, mm), entries_norm in months:
            yy_i, mm_i = int(yy), int(mm)
            existing_rows = ustore.filter_doctor_month(store_rows, doctor, yy_i, mm_i)
            diff = compute_unavailability_diff(existing_rows, entries_norm)
            if diff.get("added_count") or diff.get("removed_count") or diff.get("note_changed_count"):
                audit_todo.append((f"{yy_i}-{mm_i:02d}", diff))

            new_rows = ustore.replace_doctor_month(
                new_rows, doctor, yy_i, mm_i, entries_norm, updated_at=updated_at
            )

        try:
            _new_sha = save_store_to_github(new_rows, store_sha, message=message)

            # Read-back verification: only show "Saved" if the persisted store
            # matches the entries we intended to write for each month.
            latest_rows, latest_sha = load_store_from_github()
            for (yy, mm), entries_norm in months:
                yy_i, mm_i = int(yy), int(mm)
                persisted = ustore.filter_doctor_month(latest_rows, doctor, yy_i, mm_i)
                if _month_entries_signature(persisted) != _entries_signature_from_tuples(entries_norm):
                    raise RuntimeError(
                        "Salvataggio non verificato: i dati sul server non corrispondono a quanto inserito. "
                        "Ricarica e riprova."
                    )

            return audit_todo, (latest_sha or _new_sha)
        except Exception as e:
            last_err = e
            if _is_sha_conflict_error(e):
                # Exponential backoff + jitter to reduce repeated collisions.
                sleep_s = min(3.0, 0.35 * (2 ** attempt) + random.random() * 0.25)
                time.sleep(sleep_s)
                continue
            raise

    if last_err:
        raise last_err
    raise RuntimeError("Errore salvataggio: tentativi esauriti senza dettaglio.")

# ---------------- GitHub settings & audit log ----------------
DEFAULT_SETTINGS = {
    "unavailability_open": True,
    "max_unavailability_per_shift": 6,
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

    # optional metadata
    out["updated_at"] = str(data.get("updated_at") or "")
    out["updated_by"] = str(data.get("updated_by") or "")

    # defensive bounds
    if out["max_unavailability_per_shift"] < 0:
        out["max_unavailability_per_shift"] = 0

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
                buf.write(existing_text.strip("\n") + "\n")
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
    for _d, sh, _n in entries2:
        counts[sh] = counts.get(sh, 0) + 1

    return entries2, {"invalid_date": invalid_date, "out_of_month": out_of_month, "counts": counts}


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

    rows, sha = load_store_from_github()
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
st.title("Turni UTIC – Autogeneratore")
st.markdown(
    '<div class="small-muted">Genera il file turni del mese rispettando regole e indisponibilità. '
    'I medici possono inserire solo le <b>proprie</b> indisponibilità (privacy).</div>',
    unsafe_allow_html=True,
)

mode = st.sidebar.radio(
    "Sezione",
    ["Indisponibilità (Medico)", "Genera turni (Admin)"],
    index=0,
)

# Load default rules (for doctor list)
cfg_default = tg.load_rules(DEFAULT_RULES_PATH)
doctors_default = doctors_from_cfg(cfg_default)

# =====================================================================
#                        MEDICO – Indisponibilità
# =====================================================================
if mode == "Indisponibilità (Medico)":
    st.subheader("Indisponibilità (Medico)")
    st.write(
        "Compila le tue indisponibilità per uno o più mesi. "
        "Le indisponibilità degli altri non sono visibili."
    )

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

        doctor = st.selectbox("1) Seleziona il tuo nome", doctors_default, index=0, key="login_doctor")
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
                pin = st.text_input("2) PIN", type="password", key="login_pin", help="Il tuo PIN personale (consigliato 4 cifre)")
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
            contacts = load_doctor_contacts_from_github()
            c = contacts.get(doctor) or {}
            email = str(c.get("email") or "").strip()
            phone = str(c.get("phone") or "").strip()

            available_channels = []
            if _email_is_configured() and email:
                available_channels.append(("Email", "email", _mask_email(email)))
            if _sms_is_configured() and phone:
                available_channels.append(("SMS", "sms", _mask_phone(phone)))

            if not available_channels:
                st.error(
                    "Non posso inviare il codice di verifica: per questo medico non è configurata "
                    "una Email/numero di telefono oppure il servizio Email/SMS non è configurato."
                )
                st.info("Soluzione: aggiungi i contatti in data/doctor_contacts.yml e configura SMTP o Twilio nei secrets.")
                return

            st.caption("Per motivi di sicurezza, per impostare o resettare il PIN serve un codice inviato via Email o SMS.")

            with st.form(f"otp_request_{mode_label}", clear_on_submit=False):
                labels = [f"{lab} ({masked})" for (lab, _ch, masked) in available_channels]
                channels = [ch for (_lab, ch, _masked) in available_channels]
                idx = 0
                choice = st.selectbox("2) Dove vuoi ricevere il codice?", list(range(len(channels))), format_func=lambda i: labels[i])
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
                    code = st.text_input("3) Codice (6 cifre)", key=f"otp_code_{mode_label}")
                    new_pin = st.text_input("4) Nuovo PIN (4 cifre)", type="password", key=f"new_pin_{mode_label}")
                    new_pin2 = st.text_input("5) Conferma nuovo PIN", type="password", key=f"new_pin2_{mode_label}")
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

    sel_default = st.session_state.get("doctor_selected_months") or [(today.year, today.month)]
    sel_set = set(sel_default)

    st.subheader("3) Seleziona mese/i da compilare")
    c1, c2, c3, c4 = st.columns([1, 1.4, 1, 1])
    with c1:
        yy_sel = st.selectbox("Anno", year_options, index=0, key="doctor_year_sel")
    with c2:
        mm_sel = st.selectbox(
            "Mese",
            list(range(1, 13)),
            format_func=lambda m: f"{m:02d} - {month_names.get(m, str(m))}",
            key="doctor_month_sel",
        )
    with c3:
        add_month = st.button("Aggiungi", use_container_width=True, help="Aggiunge l’anno/mese selezionato all’elenco.")
    with c4:
        remove_month = st.button("Rimuovi", use_container_width=True, help="Rimuove l’anno/mese selezionato dall’elenco.")

    cur = (int(yy_sel), int(mm_sel))
    if add_month:
        sel_set.add(cur)
    if remove_month:
        sel_set.discard(cur)

    selected = sorted(sel_set)
    st.session_state.doctor_selected_months = selected

    st.caption("Mesi selezionati: " + ", ".join([f"{yy}-{mm:02d}" for (yy, mm) in selected]))
    if not selected:
        st.info("Aggiungi almeno un mese per iniziare.")
        st.stop()

    label_map = {(yy, mm): f"{yy}-{mm:02d}" for (yy, mm) in selected}

    # Stable baseline (snapshot) for this editing session.
    # This is what we compare against at save-time to detect a stale editor.
    cR1, cR2 = st.columns([1, 3])
    with cR1:
        refresh_baseline = st.button(
            "🔄 Ricarica dati",
            help="Ricarica l’archivio dal server (utile se qualcuno ha appena salvato).",
        )
    with cR2:
        st.caption("La sessione usa uno snapshot per evitare conflitti: al salvataggio viene sempre verificato sul server.")

    if refresh_baseline:
        # Reset baseline + editors so the UI reflects the latest server state.
        clear_doctor_baseline()
        for (yy, mm) in selected:
            st.session_state.pop(f"unav_editor_{doctor}_{yy}_{mm}", None)
        st.rerun()

    try:
        baseline = get_or_load_doctor_baseline(doctor, selected, force_reload=bool(refresh_baseline))
        store_rows = list(baseline.get("rows") or [])
        store_sha = baseline.get("sha")
        expected_signatures = dict(baseline.get("expected_signatures") or {})
    except Exception as e:
        st.error(f"Errore accesso archivio indisponibilità: {e}")
        st.stop()

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

    if not unav_open:
        st.warning("🔒 Inserimento indisponibilità temporaneamente **chiuso** dall'amministratore. Puoi solo visualizzare (non puoi salvare).")
    st.caption(
        f"Limite per medico: **max {max_per_shift}** inserimenti per ogni fascia "
        "(Mattina/Pomeriggio/Notte/Diurno/Tutto il giorno) per ogni mese."
    )

    st.divider()

    tabs = st.tabs([label_map[x] for x in selected])
    edited_by_month = {}
    normalized_entries_by_month = {}
    violations_by_month = {}
    info_by_month = {}

    for (yy, mm), tab in zip(selected, tabs):
        with tab:
            st.caption("Inserisci righe con Data + Fascia. Le righe vuote verranno ignorate.")
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

            if unav_open:
                edited = st.data_editor(
                    init,
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Data": st.column_config.DateColumn("Data", required=True),
                        "Fascia": st.column_config.SelectboxColumn("Fascia", options=FASCIA_OPTIONS, required=True),
                        "Note": st.column_config.TextColumn("Note"),
                    },
                    key=f"unav_editor_{doctor}_{yy}_{mm}",
                )
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
            over = {sh: n for sh, n in counts.items() if n > max_per_shift}
            violations_by_month[(yy, mm)] = over

            if info.get("out_of_month"):
                st.warning(
                    f"⚠️ {info['out_of_month']} righe con data fuori mese sono state ignorate "
                    f"(devono essere in {yy}-{mm:02d})."
                )
            if info.get("invalid_date"):
                st.warning(f"⚠️ {info['invalid_date']} righe hanno una data non valida e sono state ignorate.")

            st.caption(
                "Conteggi mese (per fascia): "
                + ", ".join([f"{sh} {counts.get(sh, 0)}/{max_per_shift}" for sh in FASCIA_OPTIONS])
            )

            if over:
                pretty = ", ".join([f"{sh}: {n}/{max_per_shift}" for sh, n in over.items()])
                st.error(f"Limite superato in questo mese → {pretty}. Rimuovi alcune righe prima di salvare.")

    any_over = any(bool(v) for v in (violations_by_month or {}).values())
    can_save = bool(unav_open) and (not any_over)

    c1, c2 = st.columns([1, 2])
    with c1:
        save = st.button("Salva indisponibilità", type="primary", disabled=not can_save)
    with c2:
        st.caption("Privacy: salviamo solo le righe del tuo nominativo nei mesi selezionati.")

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
            for _d, sh, _n in entries_norm:
                counts[sh] = counts.get(sh, 0) + 1
            over = {sh: n for sh, n in counts.items() if n > max_per_shift}
            if over:
                hard_viol.append(
                    f"{yy}-{mm:02d}: " + ", ".join([f"{sh} {n}/{max_per_shift}" for sh, n in over.items()])
                )

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

            st.success("Salvato ✅")
            st.rerun()
        except Exception as e:
            st.error(f"Errore salvataggio su GitHub: {e}")
            st.info(
                "Se vedi 404: (1) token senza accesso alla repo privata, "
                "(2) owner/repo/branch/path errati, (3) token non autorizzato SSO (se repo in Organization)."
            )


# =====================================================================
#                           ADMIN – Generazione
# =====================================================================
else:
    st.subheader("Generazione turni (Admin)")
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

    # ---- Admin settings: open/close unavailability + limits ----
    with st.expander("⚙️ Impostazioni indisponibilità (Admin)", expanded=True):
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

        cS1, cS2, cS3 = st.columns([1.4, 1, 2])
        with cS1:
            new_open = st.toggle(
                "Consenti ai medici di inserire/modificare indisponibilità",
                value=cur_open,
                help="Se disattivato, i medici possono solo visualizzare le proprie indisponibilità ma non salvarle.",
            )
        with cS2:
            new_max = st.number_input(
                "Max per fascia (per mese)",
                min_value=0,
                max_value=31,
                value=int(cur_max),
                step=1,
                help="Esempio: 6 significa max 6 Mattine, 6 Pomeriggi, 6 Notti, ecc. per ogni mese.",
            )
        with cS3:
            meta = ""
            if app_settings.get("updated_at"):
                meta += f"Ultimo aggiornamento: {app_settings.get('updated_at')}"
            if app_settings.get("updated_by"):
                meta += f" | da: {app_settings.get('updated_by')}"
            if meta:
                st.caption(meta)

        if st.button("Salva impostazioni indisponibilità", type="primary"):
            settings_to_save = {
                "unavailability_open": bool(new_open),
                "max_unavailability_per_shift": int(new_max),
                "updated_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                "updated_by": "admin",
            }
            try:
                save_app_settings_to_github(
                    settings_to_save,
                    app_settings_sha,
                    message=f"Update unavailability settings: open={bool(new_open)} max={int(new_max)}",
                )
                st.success("Impostazioni salvate ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Errore salvataggio impostazioni su GitHub: {e}")

    st.divider()

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

    # ---- Admin: download / inspect monthly audit log (unavailability edits) ----
    with st.expander("📜 Log inserimenti/modifiche indisponibilità (Audit)", expanded=False):
        st.write(
            "Questo log registra chi ha inserito/modificato indisponibilità, con timestamp e conteggi. "
            "È utile per tracciare le modifiche mese per mese."
        )

        cL1, cL2, cL3 = st.columns([1, 1, 2])
        with cL1:
            audit_year = st.number_input(
                "Anno log",
                min_value=2025,
                max_value=2035,
                value=int(year),
                step=1,
                key="audit_year",
            )
        with cL2:
            audit_month = st.number_input(
                "Mese log",
                min_value=1,
                max_value=12,
                value=int(month),
                step=1,
                key="audit_month",
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
            st.info(
                "Nessun audit log trovato per questo mese (nessuna modifica registrata oppure file non ancora creato)."
            )
        else:
            # Download buttons (full file)
            st.download_button(
                "⬇️ Scarica audit log (CSV)",
                data=str(audit_text).encode("utf-8"),
                file_name=f"unavailability_audit_{mk_log}.csv",
                mime="text/csv",
                key=f"dl_audit_csv_{mk_log}",
            )

            # Parse for preview + optional Excel export
            try:
                df_audit = pd.read_csv(io.StringIO(audit_text))
            except Exception as e:
                df_audit = None
                st.error("Il file audit log esiste ma non riesco a leggerlo come CSV.")
                st.code(str(e))

            if df_audit is not None and not df_audit.empty:
                try:
                    # Sort newest first if possible
                    if "ts_utc" in df_audit.columns:
                        df_audit = df_audit.sort_values("ts_utc", ascending=False)
                except Exception:
                    pass

                # Optional doctor filter for on-screen preview
                doctor_filter = "Tutti"
                if "doctor" in df_audit.columns:
                    doctors_in_log = sorted(
                        [str(x) for x in df_audit["doctor"].dropna().unique().tolist() if str(x).strip()]
                    )
                    doctor_filter = st.selectbox(
                        "Filtro medico (solo anteprima)",
                        ["Tutti"] + doctors_in_log,
                        index=0,
                        key=f"audit_filter_{mk_log}",
                    )

                df_preview = df_audit
                if doctor_filter != "Tutti" and "doctor" in df_audit.columns:
                    df_preview = df_audit[df_audit["doctor"] == doctor_filter]

                # Excel export
                try:
                    xlsx_bytes = audit_df_to_excel_bytes(df_audit, sheet_name=f"audit_{mk_log}")
                    st.download_button(
                        "⬇️ Scarica audit log (Excel)",
                        data=xlsx_bytes,
                        file_name=f"unavailability_audit_{mk_log}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_audit_xlsx_{mk_log}",
                    )
                except Exception as e:
                    st.warning(f"Esportazione Excel non disponibile: {e}")

                st.markdown("**Anteprima**")
                st.dataframe(df_preview.head(200), use_container_width=True, hide_index=True)
                st.caption("Mostro al massimo 200 righe. Per analisi completa usa il download.")

    # Step 2: Indisponibilità
    st.markdown("### 2) Indisponibilità")
    unav_mode = st.radio(
        "Fonte indisponibilità",
        ["Nessuna", "Carica file manuale", "Usa archivio (privacy)"],
        horizontal=True,
        help="Puoi caricare un file manuale, oppure usare l’archivio compilato dai medici.",
    )
    unav_upload = None
    if unav_mode == "Carica file manuale":
        unav_upload = st.file_uploader("Carica indisponibilità (xlsx/csv/tsv)", type=["xlsx", "csv", "tsv"])
    use_archive = (unav_mode == "Usa archivio (privacy)")

    # Step 3: Vincolo post-notte (carryover)
    st.markdown("### 3) Vincolo post-notte a cavallo mese")
    st.info(
        "Serve solo se qualcuno ha fatto **NOTTE l’ultimo giorno del mese precedente**: "
        "quella persona **non può lavorare il Giorno 1** del mese corrente.\n\n"
        "✅ Consigliato: carica l’**output del mese precedente**.\n"
        "🔁 Alternativa: seleziona manualmente chi ha fatto la NOTTE.",
        icon="💡",
    )

    # Admin advanced (rules/template/carryover file)
    with st.expander("⚙️ Avanzate (Regole, Template, Carryover file)", expanded=False):
        st.markdown("**Regole (solo Admin)**")
        rules_upload = st.file_uploader("Carica Regole YAML (opzionale)", type=["yml", "yaml"])
        cfg_admin, rules_path = load_rules_from_source(rules_upload)
        doctors = doctors_from_cfg(cfg_admin)

        st.markdown("**Template Excel**")
        template_upload = st.file_uploader("Carica template turni (opzionale)", type=["xlsx"])
        style_upload = st.file_uploader("Carica Style_Template.xlsx (opzionale)", type=["xlsx"])
        sheet_name = st.text_input("Nome foglio (opzionale)", value="")

        st.markdown("**Carryover – file mese precedente (opzionale)**")
        prev_out = st.file_uploader("Carica output mese precedente", type=["xlsx"], key="prev")

    # If advanced not expanded, still need cfg_admin/doctors variables
    if "cfg_admin" not in locals():
        cfg_admin, rules_path = tg.load_rules(DEFAULT_RULES_PATH), DEFAULT_RULES_PATH
        doctors = doctors_from_cfg(cfg_admin)
        template_upload = None
        style_upload = None
        sheet_name = ""

        prev_out = None

    manual_block = st.multiselect(
        "Seleziona medico/i da bloccare il Giorno 1 (se non carichi l’output precedente)",
        doctors,
        default=[],
        help="Inserisci qui chi ha fatto NOTTE l’ultimo giorno del mese precedente.",
    )

    carryover_by_month = {}
    carry_info = None

    # From file
    if prev_out is not None:
        tmp_prev = Path(tempfile.gettempdir()) / f"prev_{int(time.time())}.xlsx"
        tmp_prev.write_bytes(prev_out.getvalue())
        try:
            carry_info = tg.extract_carryover_from_output_xlsx(
                tmp_prev,
                sheet_name=sheet_name or None,
                night_col_letter="J",
                min_gap=int((cfg_admin.get("global_constraints") or {}).get("night_spacing_days_min", 5)),
            )
            carryover_by_month[mk] = carry_info
            st.success(
                f"Carryover letto: ultima data {carry_info.get('source_last_date')} | "
                f"NOTTE ultimo giorno: {carry_info.get('night_last_day_doctor')}"
            )
        except Exception as e:
            st.error(f"Errore lettura carryover: {e}")

    # Manual fallback
    if manual_block:
        carryover_by_month.setdefault(mk, {})
        carryover_by_month[mk].setdefault("blocked_day1_doctors", [])
        for d in manual_block:
            if d not in carryover_by_month[mk]["blocked_day1_doctors"]:
                carryover_by_month[mk]["blocked_day1_doctors"].append(d)

    st.divider()

    # Generate button
    generate = st.button("🚀 Genera turni", type="primary")

    if generate:
        t0 = time.time()
        status = st.status("Preparazione…", expanded=True)
        try:
            with tempfile.TemporaryDirectory() as td:
                td = Path(td)
                rules_path_use = rules_path

                status.update(label="Preparazione template…", state="running")
                if template_upload is not None:
                    template_path = td / "template.xlsx"
                    template_path.write_bytes(template_upload.getvalue())
                else:
                    # Auto template
                    if style_upload is not None:
                        style_path = td / "Style_Template.xlsx"
                        style_path.write_bytes(style_upload.getvalue())
                    else:
                        style_path = DEFAULT_STYLE_TEMPLATE if DEFAULT_STYLE_TEMPLATE.exists() else None
                    template_path = td / f"turni_{mk}.xlsx"
                    tg.create_month_template_xlsx(
                        rules_path_use,
                        int(year),
                        int(month),
                        out_path=template_path,
                        sheet_name=sheet_name or None,
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

                    store_rows_2, sha2 = load_store_from_github()
                    if sha1 and sha2 and sha2 != sha1:
                        # Archive changed during preparation: rebuild from latest.
                        rows_month = ustore.filter_month(store_rows_2, int(year), int(month))
                        xlsx_utils.build_unavailability_xlsx(rows_month, DEFAULT_UNAV_TEMPLATE, unav_path)
                        st.caption("Archivio indisponibilità aggiornato durante la preparazione: ricaricata l’ultima versione.")

                    st.caption(f"Archivio indisponibilità: {len(rows_month)} righe per {mk}")

                status.update(label="Generazione turni…", state="running")
                out_path = td / f"output_{mk}.xlsx"
                stats, log_path = tg.generate_schedule(
                    template_xlsx=template_path,
                    rules_yml=rules_path_use,
                    out_xlsx=out_path,
                    unavailability_path=unav_path,
                    sheet_name=sheet_name or None,
                    carryover_by_month=carryover_by_month if carryover_by_month else None,
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
