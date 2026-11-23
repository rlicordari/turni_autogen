# -*- coding: utf-8 -*-
"""Holiday shifts assigned in advance (draw / manual), edited from the admin page.

Stored in the data repo as data/turni_festivi.yml:

    festivi_extra:            # local holidays treated as festive (DE+HI slots)
      - {date: "2026-12-31", name: "S. Silvestro"}
    entries:                  # who works which holiday shift
      - {date: "2026-12-25", doctor: "Pugliatti", shift: "Mattina", note: "Natale"}

Shift → column: Mattina = D (DE slot on festive days), Pomeriggio = H (HI slot),
Notte = J. The night before a holiday (when it is not festive itself) is
assigned too: only its Notte. Pure logic, no Streamlit.
"""

from __future__ import annotations

import datetime as dt
import random
from typing import Iterable, List, Optional

import yaml

SHIFTS = ["Mattina", "Pomeriggio", "Notte"]
SHIFT_TO_COLUMN = {"Mattina": "D", "Pomeriggio": "H", "Notte": "J"}

HEADER = (
    "# Turni festivi pre-assegnati (sorteggio / assegnazione manuale).\n"
    "# Gestito dalla pagina Admin -> Festivi dell'app: non modificare a mano.\n"
)


def _easter(year: int) -> dt.date:
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month, day = divmod(h + l - 7 * m + 114, 31)
    return dt.date(year, month, day + 1)


def holiday_names(year: int) -> dict:
    """National holidays (with Easter Sunday) → name."""
    easter = _easter(year)
    return {
        dt.date(year, 1, 1): "Capodanno",
        dt.date(year, 1, 6): "Epifania",
        easter: "Pasqua",
        easter + dt.timedelta(days=1): "Pasquetta",
        dt.date(year, 4, 25): "Festa della Liberazione",
        dt.date(year, 5, 1): "Festa del Lavoro",
        dt.date(year, 6, 2): "Festa della Repubblica",
        dt.date(year, 8, 15): "Ferragosto",
        dt.date(year, 11, 1): "Ognissanti",
        dt.date(year, 12, 8): "Immacolata",
        dt.date(year, 12, 25): "Natale",
        dt.date(year, 12, 26): "Santo Stefano",
    }


# ── File ─────────────────────────────────────────────────────────────────────

def empty() -> dict:
    return {"festivi_extra": [], "entries": []}


def sort_entries(entries: Iterable[dict]) -> List[dict]:
    return sorted(entries, key=lambda e: (e["date"], SHIFTS.index(e["shift"]) if e["shift"] in SHIFTS else 9))


def parse(text: str) -> dict:
    data = yaml.safe_load(text or "") or {}
    if not isinstance(data, dict):
        data = {}
    extra = []
    for item in data.get("festivi_extra") or []:
        if isinstance(item, dict):
            date, name = str(item.get("date") or "").strip(), str(item.get("name") or "").strip()
        else:
            date, name = str(item or "").strip(), ""
        try:
            dt.date.fromisoformat(date)
        except ValueError:
            continue
        extra.append({"date": date, "name": name})
    entries = []
    for e in data.get("entries") or []:
        if not isinstance(e, dict):
            continue
        date, doctor = str(e.get("date") or "").strip(), str(e.get("doctor") or "").strip()
        shift = str(e.get("shift") or "").strip().capitalize()
        try:
            dt.date.fromisoformat(date)
        except ValueError:
            continue
        if doctor and shift in SHIFTS:
            entries.append({"date": date, "doctor": doctor, "shift": shift, "note": str(e.get("note") or "").strip()})
    return {"festivi_extra": sorted(extra, key=lambda x: x["date"]), "entries": sort_entries(entries)}


def to_text(data: dict) -> str:
    body = {
        "festivi_extra": [dict(x) for x in data.get("festivi_extra") or []],
        "entries": [dict(e) for e in sort_entries(data.get("entries") or [])],
    }
    return HEADER + yaml.safe_dump(body, allow_unicode=True, sort_keys=False, default_flow_style=False)


def to_solver(data: dict) -> dict:
    """Same shape as turni_generator.load_turni_festivi()."""
    return {
        "festivi_extra": [x["date"] for x in data.get("festivi_extra") or []],
        "fixed_assignments": [
            {"doctor": e["doctor"], "date": e["date"], "column": SHIFT_TO_COLUMN[e["shift"]]}
            for e in data.get("entries") or []
            if e.get("shift") in SHIFT_TO_COLUMN
        ],
    }


# ── Holiday shifts of a period ───────────────────────────────────────────────

def festive_days(start: dt.date, end: dt.date, extra: Iterable[dict]) -> dict:
    """{date: name} of the holidays from start to the day after end (the eve of that day counts)."""
    out = {}
    for year in range(start.year, end.year + 2):
        out.update(holiday_names(year))
    for x in extra or []:
        d = dt.date.fromisoformat(x["date"])
        out[d] = x.get("name") or out.get(d) or "Festivo locale"
    return {d: n for d, n in out.items() if start <= d <= end + dt.timedelta(days=1)}


def slots(start: dt.date, end: dt.date, extra: Iterable[dict]) -> List[dict]:
    """Holiday shifts to assign: Mattina/Pomeriggio/Notte of each holiday + Notte of its eve."""
    fest = festive_days(start, end, extra)
    out = []
    for d in sorted(fest):
        eve = d - dt.timedelta(days=1)
        if eve not in fest and start <= eve <= end:
            out.append({"date": eve.isoformat(), "shift": "Notte", "label": f"Pre-{fest[d]}"})
        if start <= d <= end:
            out += [{"date": d.isoformat(), "shift": sh, "label": fest[d]} for sh in SHIFTS]
    return sorted(out, key=lambda s: (s["date"], SHIFTS.index(s["shift"])))


# ── Draw ─────────────────────────────────────────────────────────────────────

def draw(
    slot_list: List[dict],
    entries: List[dict],
    *,
    day_doctors: List[str],
    night_doctors: List[str],
    only_empty: bool = True,
    seed: Optional[int] = None,
) -> List[dict]:
    """Fill the holiday shifts by draw, balancing how many each doctor gets.

    Hard: participants of that kind of shift, one shift per doctor per day, no
    morning/afternoon after that doctor's night. Then fewest holiday shifts
    overall, fewest of that kind, not next to another of their holiday shifts,
    then chance (seeded: the same seed gives the same draw).
    """
    rng = random.Random(seed)
    keys = {(s["date"], s["shift"]) for s in slot_list}
    kept = [dict(e) for e in entries if only_empty or (e["date"], e["shift"]) not in keys]
    taken = {(e["date"], e["shift"]) for e in kept}
    by_date: dict = {}
    for e in kept:
        by_date.setdefault(e["date"], {})[e["shift"]] = e["doctor"]
    total: dict = {}
    kind: dict = {}

    def _count(doc: str, shift: str) -> None:
        total[doc] = total.get(doc, 0) + 1
        k = (doc, shift == "Notte")
        kind[k] = kind.get(k, 0) + 1

    period = {s["date"] for s in slot_list}
    for e in kept:
        if e["date"] in period:
            _count(e["doctor"], e["shift"])

    for s in slot_list:
        if (s["date"], s["shift"]) in taken:
            continue
        d = dt.date.fromisoformat(s["date"])
        prev, nxt = (d - dt.timedelta(days=1)).isoformat(), (d + dt.timedelta(days=1)).isoformat()
        same_day = set((by_date.get(s["date"]) or {}).values())
        night = s["shift"] == "Notte"
        cands = []
        for doc in (night_doctors if night else day_doctors):
            if doc in same_day:
                continue
            if not night and (by_date.get(prev) or {}).get("Notte") == doc:
                continue
            if night and doc in {v for k, v in (by_date.get(nxt) or {}).items() if k != "Notte"}:
                continue
            near = doc in set((by_date.get(prev) or {}).values()) | set((by_date.get(nxt) or {}).values())
            cands.append(((total.get(doc, 0), kind.get((doc, night), 0), near, rng.random()), doc))
        if not cands:
            continue
        doc = min(cands)[1]
        kept.append({"date": s["date"], "doctor": doc, "shift": s["shift"], "note": s.get("label", "")})
        taken.add((s["date"], s["shift"]))
        by_date.setdefault(s["date"], {})[s["shift"]] = doc
        _count(doc, s["shift"])
    return sort_entries(kept)


# ── Checks ───────────────────────────────────────────────────────────────────

def check(entries: List[dict], *, active_doctors: Iterable[str]) -> List[str]:
    """Problems the admin should fix before generating."""
    active = set(active_doctors)
    entries = sort_entries(entries)
    out = []
    by_date: dict = {}
    for e in entries:
        if e["doctor"] not in active:
            out.append(f"{e['date']} {e['shift']}: {e['doctor']} non attivo nel pool medici (va sostituito).")
        by_date.setdefault(e["date"], []).append(e["doctor"])
    for date, docs in sorted(by_date.items()):
        for doc in sorted({d for d in docs if docs.count(d) > 1}):
            out.append(f"{date}: {doc} ha due turni nello stesso giorno.")
    nights = {e["date"]: e["doctor"] for e in entries if e["shift"] == "Notte"}
    for e in entries:
        if e["shift"] == "Notte":
            continue
        prev = (dt.date.fromisoformat(e["date"]) - dt.timedelta(days=1)).isoformat()
        if nights.get(prev) == e["doctor"]:
            out.append(f"{e['date']} {e['shift']}: {e['doctor']} è subito dopo la notte del {prev}.")
    return out
