# -*- coding: utf-8 -*-
"""Email receipt sent after an unavailability save.

The receipt lists what the SERVER holds after the save (read-back rows), so it
is proof of what was recorded. Free-text notes are never included: copies go
to shared mailboxes.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

LOCAL_TZ = ZoneInfo("Europe/Rome")
_MONTHS = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio",
           "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
_DOW = ["lun", "mar", "mer", "gio", "ven", "sab", "dom"]
_EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def parse_email_list(text: str) -> List[str]:
    return [p.strip() for p in re.split(r"[,;\n]", str(text or "")) if p.strip()]


def recipients(doctor_email: str, cc: Iterable[str]) -> Tuple[List[str], List[str]]:
    """(to, cc): doctor first, copies deduplicated (case-insensitive), invalid dropped."""
    seen = set()
    to: List[str] = []
    copies: List[str] = []
    for addr, bucket in [(doctor_email, to)] + [(c, copies) for c in (cc or [])]:
        addr = str(addr or "").strip()
        if not addr or not _EMAIL_RE.match(addr) or addr.lower() in seen:
            continue
        seen.add(addr.lower())
        bucket.append(addr)
    if not to:
        return copies, []
    return to, copies


def month_label(month_key: str) -> str:
    return f"{_MONTHS[int(month_key[5:7]) - 1]} {month_key[:4]}"


def _day_label(date_iso: str) -> str:
    d = dt.date.fromisoformat(str(date_iso)[:10])
    return f"{_DOW[d.weekday()]} {d:%d/%m}"


def _local_time(ts_utc: str) -> str:
    t = dt.datetime.fromisoformat(str(ts_utc).replace("Z", "+00:00"))
    if t.tzinfo is None:
        t = t.replace(tzinfo=dt.timezone.utc)
    return t.astimezone(LOCAL_TZ).strftime("%d/%m/%Y %H:%M")


def build_receipt(
    doctor: str,
    rows_by_month: Dict[str, List[dict]],
    diffs_by_month: Dict[str, dict],
    *,
    saved_at_utc: str,
    commit_sha: Optional[str],
    by_admin: bool = False,
) -> Tuple[str, str]:
    months = sorted(set(rows_by_month) | set(diffs_by_month))
    subject = f"Indisponibilità registrate – {doctor} – " + ", ".join(month_label(mk) for mk in months)
    lines = [
        "Resoconto delle indisponibilità REGISTRATE SUL SERVER (Turni UTIC/Cardiologia).",
        "Solo queste verranno usate per la generazione dei turni.",
        "",
        f"Medico: {doctor}",
        f"Salvataggio: {_local_time(saved_at_utc)} (ora italiana)"
        + (f" · codice verifica: {commit_sha[:7]}" if commit_sha else ""),
    ]
    if by_admin:
        lines.append("Modifica effettuata dall'amministratore per conto del medico.")
    for mk in months:
        rows = sorted(rows_by_month.get(mk) or [], key=lambda r: (str(r.get("date", "")), str(r.get("shift", ""))))
        diff = diffs_by_month.get(mk) or {}
        details = diff.get("details") or {}
        lines += ["", f"{month_label(mk).upper()} — {len(rows)} indisponibilità registrate"]
        if diff:
            lines.append(
                f"Variazioni rispetto a prima: +{diff.get('added_count', 0)} aggiunte, "
                f"-{diff.get('removed_count', 0)} rimosse"
            )
            for sign, key in (("+", "added"), ("-", "removed")):
                for item in details.get(key) or []:
                    lines.append(f"  {sign} {_day_label(item['date'])}  {item['shift']}")
        lines.append("Elenco completo registrato:")
        if not rows:
            lines.append("  Nessuna indisponibilità registrata per questo mese.")
        for r in rows:
            note_flag = " (con nota)" if str(r.get("note") or "").strip() else ""
            lines.append(f"  {_day_label(r['date'])}  {r.get('shift', '')}{note_flag}")
    lines += [
        "",
        "Se qualcosa non corrisponde a quanto hai inserito, contatta l'amministratore "
        "prima della generazione dei turni.",
    ]
    return subject, "\n".join(lines)
