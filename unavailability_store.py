# -*- coding: utf-8 -*-
"""Unavailability datastore utilities.

Storage format (CSV, UTF-8):
  doctor,date,shift,note,updated_at
where:
  - date is ISO YYYY-MM-DD
  - shift is one of: Mattina, Pomeriggio, Notte, Diurno, Tutto il giorno

Availability store (preferenze disponibilità) — stesso formato + campo priority:
  doctor,date,shift,note,updated_at,priority
  priority: alta | media | bassa  (default: media)
"""

from __future__ import annotations

import csv
import io
import datetime as dt
import time
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Tuple

VALID_SHIFTS = {"Mattina", "Pomeriggio", "Notte", "Diurno", "Tutto il giorno", "Ferie"}

def norm_shift(s: str) -> str:
    s0 = (s or "").strip()
    # allow some aliases
    low = s0.lower()
    if low in {"matt", "mattina"}:
        return "Mattina"
    if low in {"pom", "pomeriggio"}:
        return "Pomeriggio"
    if low in {"notte", "night"}:
        return "Notte"
    if low.startswith("diurn"):
        return "Diurno"
    if low.startswith("tutto"):
        return "Tutto il giorno"
    if low == "ferie":
        return "Ferie"
    return s0

def parse_iso_date(s: str) -> dt.date:
    return dt.date.fromisoformat(str(s).strip()[:10])

VALID_PRIORITIES = {"alta", "media", "bassa"}

def norm_priority(p: str) -> str:
    p0 = (p or "media").strip().lower()
    return p0 if p0 in VALID_PRIORITIES else "media"

def load_store(csv_text: str) -> List[Dict[str, str]]:
    if not csv_text.strip():
        return []
    f = io.StringIO(csv_text)
    rdr = csv.DictReader(f)
    out: List[Dict[str, str]] = []
    for row in rdr:
        if not row:
            continue
        doctor = (row.get("doctor") or "").strip()
        date = (row.get("date") or "").strip()
        shift = (row.get("shift") or "").strip()
        if not doctor or not date or not shift:
            continue
        out.append({
            "doctor": doctor,
            "date": date[:10],
            "shift": norm_shift(shift),
            "note": (row.get("note") or "").strip(),
            "updated_at": (row.get("updated_at") or "").strip(),
            "priority": norm_priority(row.get("priority", "media")),
        })
    return out

def to_csv(rows: List[Dict[str, str]]) -> str:
    buf = io.StringIO()
    fieldnames = ["doctor", "date", "shift", "note", "updated_at", "priority"]
    wr = csv.DictWriter(buf, fieldnames=fieldnames)
    wr.writeheader()
    for r in rows:
        wr.writerow({
            "doctor": r.get("doctor",""),
            "date": r.get("date","")[:10],
            "shift": norm_shift(r.get("shift","")),
            "note": r.get("note",""),
            "updated_at": r.get("updated_at",""),
            "priority": norm_priority(r.get("priority","media")),
        })
    return buf.getvalue()

def filter_doctor_month(rows: List[Dict[str, str]], doctor: str, year: int, month: int) -> List[Dict[str, str]]:
    out=[]
    for r in rows:
        if (r.get("doctor") or "") != doctor:
            continue
        try:
            d = parse_iso_date(r.get("date",""))
        except Exception:
            continue
        if d.year==year and d.month==month:
            out.append(r)
    return out

def filter_month(rows: List[Dict[str, str]], year: int, month: int) -> List[Dict[str, str]]:
    out=[]
    for r in rows:
        try:
            d = parse_iso_date(r.get("date",""))
        except Exception:
            continue
        if d.year==year and d.month==month:
            out.append(r)
    return out

def replace_doctor_month(
    rows: List[Dict[str, str]],
    doctor: str,
    year: int,
    month: int,
    new_entries: Iterable[Tuple[dt.date, str, str]],
    updated_at: Optional[str] = None,
) -> List[Dict[str, str]]:
    """Replace all entries for doctor+month with new_entries."""
    updated_at = updated_at or dt.datetime.now(dt.timezone.utc).isoformat()
    doctor = doctor.strip()
    kept=[]
    for r in rows:
        if (r.get("doctor") or "") != doctor:
            kept.append(r); continue
        try:
            d = parse_iso_date(r.get("date",""))
        except Exception:
            continue
        if d.year==year and d.month==month:
            continue  # drop
        kept.append(r)

    for entry in new_entries:
        # entry can be (date, shift, note) or (date, shift, note, priority)
        if len(entry) == 3:
            d, shift, note = entry
            priority = "media"
        else:
            d, shift, note, priority = entry[:4]
        if not isinstance(d, dt.date):
            continue
        sh = norm_shift(shift)
        if sh not in VALID_SHIFTS:
            continue
        kept.append({
            "doctor": doctor,
            "date": d.isoformat(),
            "shift": sh,
            "note": (note or "").strip(),
            "updated_at": updated_at,
            "priority": norm_priority(priority),
        })

    # de-duplicate by (doctor,date,shift) keep latest
    dedup: Dict[Tuple[str,str,str], Dict[str,str]] = {}
    for r in kept:
        k=(r.get("doctor",""), r.get("date","")[:10], norm_shift(r.get("shift","")))
        prev=dedup.get(k)
        if not prev:
            dedup[k]=r
        else:
            # choose lexicographically larger updated_at as "newer" if iso format
            if (r.get("updated_at","") or "") >= (prev.get("updated_at","") or ""):
                dedup[k]=r
    return list(dedup.values())


# ── Month signatures & diff ──────────────────────────────────────────────────
# A signature is the sorted set of (date_iso, shift, note) of one doctor-month:
# two editors/files hold the same data iff their signatures are equal.

Signature = List[Tuple[str, str, str]]


def entries_signature(entries: Iterable[Tuple]) -> Signature:
    sig = set()
    for entry in entries or []:
        d, sh, note = entry[0], entry[1], entry[2] if len(entry) > 2 else ""
        if not isinstance(d, dt.date):
            continue
        sh2 = norm_shift(sh)
        if sh2:
            sig.add((d.isoformat(), sh2, str(note or "")))
    return sorted(sig)


def month_signature(rows: List[Dict[str, str]], doctor: str, year: int, month: int) -> Signature:
    return rows_signature(filter_doctor_month(rows, doctor, year, month))


def normalize_signature(sig: Iterable) -> Signature:
    """Signatures read back from JSON come as lists: make them comparable."""
    return sorted(tuple(str(x) for x in item) for item in (sig or []))


def compute_month_diff(existing_rows: List[Dict[str, str]], new_entries: Iterable[Tuple]) -> dict:
    """Diff (date, shift) between stored rows and edited entries; notes are not logged."""
    before = {(d, sh): note for d, sh, note in rows_signature(existing_rows)}
    after = {(d, sh): note for d, sh, note in entries_signature(new_entries)}
    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    note_changed = sorted(k for k in set(before) & set(after) if before[k] != after[k])
    return {
        "before_count": len(before),
        "after_count": len(after),
        "added_count": len(added),
        "removed_count": len(removed),
        "note_changed_count": len(note_changed),
        "details": {
            "added": [{"date": d, "shift": sh} for d, sh in added],
            "removed": [{"date": d, "shift": sh} for d, sh in removed],
            "note_changed": [{"date": d, "shift": sh} for d, sh in note_changed],
        },
    }


def rows_signature(rows: List[Dict[str, str]]) -> Signature:
    """Signature of already-filtered rows (one doctor-month)."""
    sig = set()
    for r in rows or []:
        try:
            d_iso = parse_iso_date(r.get("date", "")).isoformat()
        except Exception:
            continue
        sh = norm_shift(r.get("shift", ""))
        if sh:
            sig.add((d_iso, sh, str(r.get("note", "") or "")))
    return sorted(sig)


# ── Concurrency-safe save ────────────────────────────────────────────────────

class ShaConflictError(RuntimeError):
    """The file changed on the server after it was read (optimistic lock failed)."""


class MonthConflictError(RuntimeError):
    """The month was changed by another session after the editor was opened."""

    def __init__(self, month_key: str):
        super().__init__(
            f"Le indisponibilità di {month_key} sono state modificate da un'altra sessione "
            "o da un altro dispositivo dopo che hai aperto questa pagina."
        )
        self.month_key = month_key


class SaveNotVerifiedError(RuntimeError):
    """The write succeeded but the server never returned the saved data."""


@dataclass
class SaveOutcome:
    changed: bool
    diffs: Dict[str, dict] = field(default_factory=dict)
    verified_rows: List[Dict[str, str]] = field(default_factory=list)
    sha: Optional[str] = None
    commit_sha: Optional[str] = None


def save_doctor_months(
    *,
    load_fn: Callable[[], Tuple[List[Dict[str, str]], Optional[str]]],
    save_fn: Callable[[List[Dict[str, str]], Optional[str]], dict],
    doctor: str,
    entries_by_month: Dict[Tuple[int, int], List[Tuple]],
    base_signatures: Dict[Tuple[int, int], Iterable],
    updated_at: str,
    max_retries: int = 6,
    verify_reads: int = 6,
    sleep_fn: Callable[[float], None] = time.sleep,
    is_conflict: Optional[Callable[[Exception], bool]] = None,
) -> SaveOutcome:
    """Replace whole doctor-months with the edited entries, safely.

    - A month that another session changed after the editor was opened
      (server signature != base signature) raises MonthConflictError instead of
      being overwritten: blindly re-applying the editor would erase those days.
    - Months already equal to the edited entries are not rewritten, so a double
      tap on Save is a no-op.
    - SHA conflicts (file changed, but not in the edited months) are retried on
      fresh data; the read-back after the write is retried without rewriting.

    load_fn() -> (rows, sha); save_fn(rows, sha) -> {"content_sha", "commit_sha"}.
    """
    is_conflict = is_conflict or (lambda e: isinstance(e, ShaConflictError))
    months = sorted(entries_by_month.items())
    bases = {k: normalize_signature(v) for k, v in (base_signatures or {}).items()}
    attempted_diffs: Dict[str, dict] = {}

    for attempt in range(max_retries):
        rows, sha = load_fn()
        doctor_rows = [r for r in rows if (r.get("doctor") or "") == doctor]
        new_rows = list(doctor_rows)
        diffs: Dict[str, dict] = {}
        for (yy, mm), entries in months:
            mk = f"{int(yy):04d}-{int(mm):02d}"
            current = month_signature(doctor_rows, doctor, int(yy), int(mm))
            if current == entries_signature(entries):
                continue
            base = bases.get((yy, mm))
            if base is not None and current != base:
                raise MonthConflictError(mk)
            diffs[mk] = compute_month_diff(filter_doctor_month(doctor_rows, doctor, int(yy), int(mm)), entries)
            new_rows = replace_doctor_month(new_rows, doctor, int(yy), int(mm), entries, updated_at=updated_at)

        if not diffs:
            # Already as wanted. If an earlier attempt of THIS save had changes, its
            # write landed but the response was lost (timeout → retry → conflict).
            return SaveOutcome(
                changed=bool(attempted_diffs),
                diffs=attempted_diffs,
                verified_rows=doctor_rows,
                sha=sha,
            )
        if not attempted_diffs:
            attempted_diffs = diffs

        try:
            result = save_fn(new_rows, sha) or {}
        except Exception as e:
            if is_conflict(e) and attempt < max_retries - 1:
                sleep_fn(min(3.0, 0.35 * (2 ** attempt)))
                continue
            raise

        for i in range(verify_reads):
            v_rows, v_sha = load_fn()
            v_doctor_rows = [r for r in v_rows if (r.get("doctor") or "") == doctor]
            if all(
                month_signature(v_doctor_rows, doctor, int(yy), int(mm)) == entries_signature(entries)
                for (yy, mm), entries in months
            ):
                return SaveOutcome(
                    changed=True,
                    diffs=diffs,
                    verified_rows=v_doctor_rows,
                    sha=v_sha or result.get("content_sha"),
                    commit_sha=result.get("commit_sha"),
                )
            sleep_fn(min(2.0, 0.5 * (i + 1)))
        raise SaveNotVerifiedError(
            "Salvataggio non verificato: i dati sul server non corrispondono a quanto inserito. "
            "Ricarica e riprova."
        )

    raise ShaConflictError("Salvataggio non riuscito: troppi aggiornamenti concorrenti, riprova.")
