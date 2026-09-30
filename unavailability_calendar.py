# -*- coding: utf-8 -*-
"""Calendar editor for the doctor page: pure logic (no Streamlit).

The calendar component shows one month; each tap produces an event that is
applied here to the editor rows kept in the page session:

- unavailability rows: {"id", "Data": date, "Fascia", "Note"}
- preference rows:     {"id", "Data": date, "Fascia", "Priorita", "Note"}

Events:
- {"type": "set_day", "kind": "unav"|"pref", "day": d, "shifts": [...], "note": str, "priority": str}
  replaces that day's rows of that kind ("Ferie" and "Tutto il giorno" are exclusive);
- {"type": "range", "from": d1, "to": d2}: Ferie on every day of the range.
Invalid events change nothing.

The component value carries every edit not yet acknowledged by the page
(see apply_value): Streamlit may merge two quick taps into one rerun, and a
tap must never be lost.
"""

from __future__ import annotations

import calendar
import datetime as dt
import uuid
from typing import Callable

import unavailability_store as ustore

UNAV_SHIFTS = ["Mattina", "Pomeriggio", "Notte", "Diurno", "Tutto il giorno", "Ferie"]
PREF_SHIFTS = ["Mattina", "Pomeriggio", "Notte", "Diurno", "Tutto il giorno"]
EXCLUSIVE = ["Ferie", "Tutto il giorno"]   # first match wins
PRIORITIES = ["alta", "media", "bassa"]
MONTHS = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio",
          "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]


def _day(year: int, month: int, value) -> dt.date | None:
    try:
        d = int(value)
    except (TypeError, ValueError):
        return None
    if not 1 <= d <= calendar.monthrange(year, month)[1]:
        return None
    return dt.date(year, month, d)


def _clean_shifts(kind: str, shifts) -> list[str]:
    allowed = UNAV_SHIFTS if kind == "unav" else PREF_SHIFTS
    chosen = [s for s in allowed if s in set(shifts or [])]
    if kind == "unav":
        for exclusive in EXCLUSIVE:
            if exclusive in chosen:
                return [exclusive]
    return chosen


def apply_event(
    unav_rows: list[dict],
    pref_rows: list[dict],
    event: dict,
    *,
    year: int,
    month: int,
    new_id: Callable[[], str] = lambda: str(uuid.uuid4()),
) -> tuple[list[dict], list[dict]]:
    unav_rows = list(unav_rows or [])
    pref_rows = list(pref_rows or [])
    kind_type = (event or {}).get("type")

    if kind_type == "set_day":
        kind = event.get("kind")
        day = _day(year, month, event.get("day"))
        if kind not in ("unav", "pref") or day is None:
            return unav_rows, pref_rows
        shifts = _clean_shifts(kind, event.get("shifts"))
        note = str(event.get("note") or "").strip()
        if kind == "unav":
            kept = [r for r in unav_rows if r.get("Data") != day]
            return kept + [{"id": new_id(), "Data": day, "Fascia": s, "Note": note} for s in shifts], pref_rows
        priority = event.get("priority") if event.get("priority") in PRIORITIES else "media"
        kept = [r for r in pref_rows if r.get("Data") != day]
        return unav_rows, kept + [
            {"id": new_id(), "Data": day, "Fascia": s, "Priorita": priority, "Note": note} for s in shifts
        ]

    if kind_type == "range":
        start = _day(year, month, event.get("from"))
        end = _day(year, month, event.get("to"))
        if start is None or end is None or end < start:
            return unav_rows, pref_rows
        days = {start + dt.timedelta(days=i) for i in range((end - start).days + 1)}
        kept = [r for r in unav_rows if r.get("Data") not in days]
        added = [{"id": new_id(), "Data": d, "Fascia": "Ferie", "Note": ""} for d in sorted(days)]
        return kept + added, pref_rows

    return unav_rows, pref_rows


def _int(value) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def apply_value(
    unav_rows: list[dict],
    pref_rows: list[dict],
    value,
    *,
    year: int,
    month: int,
    ack: dict,
    new_id: Callable[[], str] = lambda: str(uuid.uuid4()),
) -> tuple[list[dict], list[dict], dict, tuple[int, int] | None]:
    """Applies the component's unacknowledged edits, in order.

    value = {"cid": instance id, "edits": [{"seq", "year", "month", **event}],
             "nav": {"seq", "year", "month"} | None}
    ack   = {"cid", "seq"}: last edit applied for that component instance.
    Returns (unav_rows, pref_rows, new_ack, month_to_show or None).
    """
    unav_rows, pref_rows = list(unav_rows or []), list(pref_rows or [])
    ack = dict(ack or {})
    if not isinstance(value, dict) or not value.get("cid"):
        return unav_rows, pref_rows, ack, None
    cid = str(value["cid"])
    last = (_int(ack.get("seq")) or 0) if ack.get("cid") == cid else 0
    top = last

    edits = value.get("edits") if isinstance(value.get("edits"), list) else []
    fresh = [e for e in edits if isinstance(e, dict) and (_int(e.get("seq")) or 0) > last]
    for e in sorted(fresh, key=lambda e: _int(e.get("seq"))):
        top = max(top, _int(e["seq"]))
        if (_int(e.get("year")), _int(e.get("month"))) == (year, month):
            unav_rows, pref_rows = apply_event(unav_rows, pref_rows, e, year=year, month=month, new_id=new_id)

    target = None
    nav = value.get("nav")
    if isinstance(nav, dict) and (_int(nav.get("seq")) or 0) > last:
        top = max(top, _int(nav["seq"]))
        y, m = _int(nav.get("year")), _int(nav.get("month"))
        if y and m and 1 <= m <= 12 and (y, m) != (year, month):
            target = (y, m)

    return unav_rows, pref_rows, ({"cid": cid, "seq": top} if top > last or ack.get("cid") != cid else ack), target


def _editor_day_items(rows: list[dict], day: dt.date, order: list[str], with_priority: bool) -> list[dict]:
    items = []
    for r in rows:
        if r.get("Data") != day:
            continue
        item = {"shift": str(r.get("Fascia") or ""), "note": str(r.get("Note") or "")}
        if with_priority:
            item["priority"] = ustore.norm_priority(r.get("Priorita", "media"))
        items.append(item)
    return sorted(items, key=lambda it: order.index(it["shift"]) if it["shift"] in order else 99)


def _stored_day_items(rows: list[dict], day: dt.date, order: list[str], with_priority: bool) -> list[dict]:
    items = []
    for r in rows:
        if str(r.get("date", ""))[:10] != day.isoformat():
            continue
        item = {"shift": ustore.norm_shift(r.get("shift", "")), "note": str(r.get("note") or "")}
        if with_priority:
            item["priority"] = ustore.norm_priority(r.get("priority", "media"))
        items.append(item)
    return sorted(items, key=lambda it: order.index(it["shift"]) if it["shift"] in order else 99)


def calendar_payload(
    unav_rows: list[dict],
    pref_rows: list[dict],
    sent_unav_rows: list[dict],
    sent_pref_rows: list[dict],
    *,
    year: int,
    month: int,
    limits: dict,
    nav: dict | None = None,
    ack: dict | None = None,
) -> dict:
    """Data for the calendar component: one entry per day, pending = differs from server."""
    from turni_generator import italy_public_holidays

    holidays = italy_public_holidays(year)
    days = []
    for d in range(1, calendar.monthrange(year, month)[1] + 1):
        day = dt.date(year, month, d)
        unav = _editor_day_items(unav_rows, day, UNAV_SHIFTS, False)
        pref = _editor_day_items(pref_rows, day, PREF_SHIFTS, True)
        sent_unav = _stored_day_items(sent_unav_rows, day, UNAV_SHIFTS, False)
        sent_pref = _stored_day_items(sent_pref_rows, day, PREF_SHIFTS, True)
        days.append({
            "day": d,
            "weekday": day.weekday(),
            "festive": day.weekday() == 6 or day in holidays,
            "unav": unav,
            "pref": pref,
            "pending": unav != sent_unav or pref != sent_pref,
        })
    return {
        "year": year,
        "month": month,
        "label": f"{MONTHS[month - 1]} {year}",
        "first_weekday": dt.date(year, month, 1).weekday(),
        "days": days,
        "limits": dict(limits or {}),
        "nav": dict(nav or {"prev": True, "next": True}),
        "ack": dict(ack or {}),
        "unav_shifts": UNAV_SHIFTS,
        "pref_shifts": PREF_SHIFTS,
    }
