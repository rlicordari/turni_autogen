# -*- coding: utf-8 -*-
"""SMTP settings editable from the admin panel (pure logic, no Streamlit).

The data repo is public, so the password (a Google app password) is stored
only encrypted: Fernet, with a key derived from a secret the app already has
(the GitHub token in Secrets). If that token changes, the stored password can
no longer be read: the app falls back to Secrets [smtp] and asks the admin to
enter it again.

Stored file (data/email_settings.json):
    {"username", "from", "host", "port", "starttls", "password_enc",
     "password_set_at", "updated_at", "updated_by"}
Empty fields fall back to Secrets [smtp]; host/port default to Gmail.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import re
from typing import Optional

from cryptography.fernet import Fernet, InvalidToken

GMAIL = {"host": "smtp.gmail.com", "port": 587, "starttls": True}
_KEY_INFO = b"turni-autogen/smtp-password/v1"


def _fernet(secret: str) -> Fernet:
    digest = hmac.new(str(secret or "").encode("utf-8"), _KEY_INFO, hashlib.sha256).digest()
    return Fernet(base64.urlsafe_b64encode(digest))


def encrypt_password(password: str, secret: str) -> str:
    return _fernet(secret).encrypt(str(password).encode("utf-8")).decode("ascii")


def decrypt_password(token: str, secret: str) -> Optional[str]:
    try:
        return _fernet(secret).decrypt(str(token or "").encode("ascii")).decode("utf-8")
    except (InvalidToken, ValueError, TypeError):
        return None


def _clean_password(password: str) -> str:
    # Google shows app passwords in groups of four ("abcd efgh ijkl mnop").
    password = str(password or "").strip()
    if re.fullmatch(r"[A-Za-z]{4}(\s*[A-Za-z]{4}){3}", password):
        return re.sub(r"\s+", "", password)
    return password


def updated_settings(
    previous: Optional[dict],
    *,
    username: str,
    from_addr: str,
    host: str,
    port: int,
    starttls: bool,
    new_password: str,
    secret: str,
    now: str,
    by: str = "admin",
) -> dict:
    """New stored settings; an empty password keeps the previous one."""
    prev = dict(previous or {})
    host = str(host or "").strip() or GMAIL["host"]
    out = {
        "username": str(username or "").strip(),
        "from": str(from_addr or "").strip(),
        "host": host,
        "port": int(port or 0) or GMAIL["port"],
        "starttls": bool(starttls),
        "password_enc": prev.get("password_enc", ""),
        "password_set_at": prev.get("password_set_at", ""),
        "updated_at": now,
        "updated_by": by,
    }
    password = _clean_password(new_password)
    if password:
        out["password_enc"] = encrypt_password(password, secret)
        out["password_set_at"] = now
    return out


def effective_smtp(secrets_cfg: dict, stored: Optional[dict], secret: str) -> tuple[dict, list[str]]:
    """Secrets [smtp] overridden by the admin panel's settings, plus warnings for the admin."""
    cfg = dict(secrets_cfg or {})
    if not stored:
        return cfg, []
    warnings = []
    for key in ("username", "host", "port", "starttls"):
        if stored.get(key) not in (None, ""):
            cfg[key] = stored[key]
    cfg["from"] = stored.get("from") or stored.get("username") or cfg.get("from", "")
    if stored.get("password_enc"):
        password = decrypt_password(stored["password_enc"], secret)
        if password is None:
            warnings.append(
                "La password salvata nel pannello non è più leggibile (è cambiato il token GitHub): "
                "reinseriscila. Nel frattempo si usano i Secrets."
            )
        else:
            cfg["password"] = password
    return cfg, warnings


def public_view(stored: Optional[dict]) -> dict:
    """What the admin panel may show: never the password, not even encrypted."""
    view = {k: v for k, v in (stored or {}).items() if k != "password_enc"}
    view["password_set"] = bool((stored or {}).get("password_enc"))
    return view
