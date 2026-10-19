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


def _diff_lines(diff: dict) -> List[str]:
    lines = [
        f"Variazioni rispetto a prima: +{diff.get('added_count', 0)} aggiunte, "
        f"-{diff.get('removed_count', 0)} rimosse"
        + (f", {diff['note_changed_count']} modificate" if diff.get("note_changed_count") else "")
    ]
    details = diff.get("details") or {}
    for sign, key in (("+", "added"), ("-", "removed")):
        for item in details.get(key) or []:
            lines.append(f"  {sign} {_day_label(item['date'])}  {item['shift']}")
    return lines


_OF = {
    "Indisponibilità": "delle indisponibilità",
    "Preferenze": "delle preferenze",
    "Indisponibilità e preferenze": "di indisponibilità e preferenze",
}


def _sorted_rows(rows) -> List[dict]:
    return sorted(rows or [], key=lambda r: (str(r.get("date", "")), str(r.get("shift", ""))))


def build_receipt(
    doctor: str,
    rows_by_month: Dict[str, List[dict]],
    diffs_by_month: Dict[str, dict],
    *,
    saved_at_utc: str,
    commit_sha: Optional[str],
    by_admin: bool = False,
    pref_rows_by_month: Optional[Dict[str, List[dict]]] = None,
    pref_diffs_by_month: Optional[Dict[str, dict]] = None,
) -> Tuple[str, str]:
    pref_rows_by_month = pref_rows_by_month or {}
    pref_diffs_by_month = pref_diffs_by_month or {}
    unav_months = set(rows_by_month) | set(diffs_by_month)
    pref_months = set(pref_rows_by_month) | set(pref_diffs_by_month)
    months = sorted(unav_months | pref_months)
    what = (
        "Indisponibilità e preferenze" if unav_months and pref_months
        else "Preferenze" if pref_months else "Indisponibilità"
    )
    subject = f"{what} registrate – {doctor} – " + ", ".join(month_label(mk) for mk in months)
    lines = [
        f"Resoconto {_OF[what]} REGISTRATE SUL SERVER (Turni UTIC/Cardiologia).",
        "Solo queste verranno usate per la generazione dei turni.",
        "",
        f"Medico: {doctor}",
        f"Salvataggio: {_local_time(saved_at_utc)} (ora italiana)"
        + (f" · codice verifica: {commit_sha[:7]}" if commit_sha else ""),
    ]
    if by_admin:
        lines.append("Modifica effettuata dall'amministratore per conto del medico.")
    for mk in months:
        if mk in unav_months:
            rows = _sorted_rows(rows_by_month.get(mk))
            diff = diffs_by_month.get(mk) or {}
            lines += ["", f"{month_label(mk).upper()} — {len(rows)} indisponibilità registrate"]
            if diff:
                lines += _diff_lines(diff)
            lines.append("Elenco completo registrato:")
            if not rows:
                lines.append("  Nessuna indisponibilità registrata per questo mese.")
            for r in rows:
                note_flag = " (con nota)" if str(r.get("note") or "").strip() else ""
                lines.append(f"  {_day_label(r['date'])}  {r.get('shift', '')}{note_flag}")
        if mk in pref_months:
            rows = _sorted_rows(pref_rows_by_month.get(mk))
            diff = pref_diffs_by_month.get(mk) or {}
            lines += ["", f"PREFERENZE {month_label(mk).upper()} — {len(rows)} registrate"]
            if diff:
                lines += _diff_lines(diff)
            lines.append("Elenco completo registrato:")
            if not rows:
                lines.append("  Nessuna preferenza registrata per questo mese.")
            for r in rows:
                note_flag = " (con nota)" if str(r.get("note") or "").strip() else ""
                lines.append(
                    f"  {_day_label(r['date'])}  {r.get('shift', '')} (priorità {r.get('priority') or 'media'}){note_flag}"
                )
    lines += [
        "",
        "Se qualcosa non corrisponde a quanto hai inserito, contatta l'amministratore "
        "prima della generazione dei turni.",
    ]
    return subject, "\n".join(lines)
