# -*- coding: utf-8 -*-
"""Doctor contacts file (data/doctor_contacts.yml): one email per doctor.

The same address receives the PIN codes (OTP) and the receipt mail after every
save, so it is edited in one place only (admin panel, pool → Medici).
"""

from __future__ import annotations

import re

import yaml

_EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def is_valid_email(addr: str) -> bool:
    return bool(_EMAIL_RE.match(str(addr or "").strip()))


def set_email(text: str, doctor: str, email: str) -> str:
    """New YAML text with the doctor's email set (empty = removed); other data kept.

    Raises ValueError for an invalid address.
    """
    email = str(email or "").strip()
    if email and not is_valid_email(email):
        raise ValueError(f"Indirizzo email non valido: {email}")
    data = yaml.safe_load(text or "") or {}
    if not isinstance(data, dict):
        data = {}
    entry = dict(data.get(doctor) or {}) if isinstance(data.get(doctor), dict) else {}
    if email:
        entry["email"] = email
    else:
        entry.pop("email", None)
    if entry:
        data[doctor] = entry
    else:
        data.pop(doctor, None)
    return yaml.safe_dump({k: data[k] for k in sorted(data)}, allow_unicode=True, default_flow_style=False, sort_keys=False)
