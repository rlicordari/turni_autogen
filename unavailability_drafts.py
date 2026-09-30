# -*- coding: utf-8 -*-
"""Drafts of unavailability edits not yet sent with "Salva".

One JSON file per doctor on GitHub (data/unavailability_drafts/draft_<slug>.json):

    {"schema_version": 1, "doctor": "...",
     "months": {"2026-10": {"entries": [["2026-10-07", "Ferie", ""], ...],
                             "base_signature": [[date, shift, note], ...],
                             "updated_at": "...Z", "session_id": "..."}}}

`base_signature` is the server data the editor started from: if the server
changed after it, the draft must not be resumed silently.

Drafts are never used by the solver: only the official per-doctor CSV is.
"""

from __future__ import annotations

import copy
import datetime as dt
import json
from typing import Iterable, List, Optional, Tuple

import unavailability_store as ustore

SCHEMA_VERSION = 1
DRAFTS_DIR_DEFAULT = "data/unavailability_drafts"

Entry = Tuple[dt.date, str, str]


def empty_draft(doctor: str) -> dict:
    return {"schema_version": SCHEMA_VERSION, "doctor": doctor, "months": {}}


def to_text(draft: dict) -> str:
    return json.dumps(draft, ensure_ascii=False, indent=2, sort_keys=True)


def from_text(text: str) -> dict:
    data = json.loads(text) if str(text or "").strip() else {}
    if not isinstance(data, dict):
        data = {}
    months = data.get("months") if isinstance(data.get("months"), dict) else {}
    return {"schema_version": SCHEMA_VERSION, "doctor": str(data.get("doctor") or ""), "months": months}


def set_month(
    draft: dict,
    month_key: str,
    *,
    entries: Iterable[Entry],
    base_signature: Iterable,
    updated_at: str,
    session_id: str,
) -> dict:
    out = copy.deepcopy(draft)
    out.setdefault("months", {})[month_key] = {
        "entries": [list(item) for item in ustore.entries_signature(entries)],
        "base_signature": [list(item) for item in ustore.normalize_signature(base_signature)],
        "updated_at": updated_at,
        "session_id": session_id,
    }
    return out


def remove_month(draft: dict, month_key: str) -> dict:
    out = copy.deepcopy(draft)
    out.setdefault("months", {}).pop(month_key, None)
    return out


def month_entries(draft: dict, month_key: str) -> List[Entry]:
    month = (draft.get("months") or {}).get(month_key) or {}
    out: List[Entry] = []
    for item in month.get("entries") or []:
        try:
            out.append((dt.date.fromisoformat(str(item[0])[:10]), str(item[1]), str(item[2] if len(item) > 2 else "")))
        except Exception:
            continue
    return out


def editor_start(draft: Optional[dict], month_key: str, official_signature: Iterable) -> dict:
    """How the month editor should start.

    - "official": no draft for the month, or the draft equals the server data;
    - "resume": unsent draft built on the current server data → reload it;
    - "offer": unsent draft, but the server changed after it → let the doctor choose.
    """
    official = ustore.normalize_signature(official_signature)
    month = ((draft or {}).get("months") or {}).get(month_key)
    if not month:
        return {"action": "official"}
    entries = month_entries(draft, month_key)
    if ustore.entries_signature(entries) == official:
        return {"action": "official"}
    base = ustore.normalize_signature(month.get("base_signature") or [])
    return {
        "action": "resume" if base == official else "offer",
        "entries": entries,
        "updated_at": month.get("updated_at") or "",
    }


class MemoryDraftStore:
    """Drafts kept in the app process, copied to GitHub at most every N seconds.

    Autosave writes here (instant, no API call). The GitHub copy (for restarts
    and as evidence) is a read-modify-write MERGE of the months changed in
    memory, so a failed initial load can never wipe a draft already on GitHub.
    """

    def __init__(
        self,
        *,
        load_fn,
        persist_fn,
        clock=None,
        persist_min_interval: float = 120.0,
    ):
        import threading
        import time as _time

        self._load_fn = load_fn          # doctor -> draft
        self._persist_fn = persist_fn    # (doctor, apply_fn(draft) -> draft) -> None
        self._clock = clock or _time.time
        self._interval = persist_min_interval
        self._lock = threading.RLock()
        self._drafts: dict = {}
        self._loaded: set = set()
        self._dirty: dict = {}           # doctor -> {month_key: month_dict | None (removed)}
        self._last_persist: dict = {}

    def get(self, doctor: str) -> dict:
        with self._lock:
            if doctor not in self._loaded:
                try:
                    server = self._load_fn(doctor)
                    self._loaded.add(doctor)
                except Exception:
                    server = empty_draft(doctor)
                mem = self._drafts.get(doctor) or empty_draft(doctor)
                merged = copy.deepcopy(server)
                merged["doctor"] = doctor
                merged.setdefault("months", {}).update(mem.get("months") or {})
                for mk, month in (self._dirty.get(doctor) or {}).items():
                    if month is None:
                        merged["months"].pop(mk, None)
                self._drafts[doctor] = merged
            return copy.deepcopy(self._drafts[doctor])

    def all_drafts(self) -> dict:
        with self._lock:
            return {d: copy.deepcopy(v) for d, v in self._drafts.items()}

    def set_month(self, doctor, month_key, *, entries, base_signature, updated_at, session_id) -> None:
        with self._lock:
            draft = self._drafts.get(doctor) or empty_draft(doctor)
            draft = set_month(draft, month_key, entries=entries, base_signature=base_signature,
                              updated_at=updated_at, session_id=session_id)
            self._drafts[doctor] = draft
            self._dirty.setdefault(doctor, {})[month_key] = copy.deepcopy(draft["months"][month_key])

    def remove_month(self, doctor, month_key) -> None:
        with self._lock:
            draft = self._drafts.get(doctor) or empty_draft(doctor)
            self._drafts[doctor] = remove_month(draft, month_key)
            self._dirty.setdefault(doctor, {})[month_key] = None

    def flush(self, doctor: str, *, force: bool = False) -> bool:
        """Copy dirty months to GitHub. Returns True if nothing is left to copy."""
        with self._lock:
            ops = dict(self._dirty.get(doctor) or {})
            if not ops:
                return True
            if not force and self._clock() - self._last_persist.get(doctor, 0.0) < self._interval:
                return False

        def _apply(server_draft: dict) -> dict:
            out = copy.deepcopy(server_draft)
            out["doctor"] = doctor
            months = out.setdefault("months", {})
            for mk, month in ops.items():
                if month is None:
                    months.pop(mk, None)
                else:
                    months[mk] = copy.deepcopy(month)
            return out

        try:
            self._persist_fn(doctor, _apply)
        except Exception:
            return False
        with self._lock:
            dirty = self._dirty.get(doctor) or {}
            for mk, month in ops.items():
                if dirty.get(mk, "missing") == month:
                    dirty.pop(mk, None)
            self._last_persist[doctor] = self._clock()
            return not dirty


def pending_summary(draft: Optional[dict], official_rows: List[dict], month_keys: Iterable[str]) -> List[dict]:
    """Months where the draft still differs from the official server data."""
    doctor = (draft or {}).get("doctor") or ""
    out = []
    for mk in month_keys:
        month = ((draft or {}).get("months") or {}).get(mk)
        if not month:
            continue
        yy, mm = int(mk[:4]), int(mk[5:7])
        entries = month_entries(draft, mk)
        official = ustore.filter_doctor_month(official_rows, doctor, yy, mm)
        if ustore.entries_signature(entries) == ustore.rows_signature(official):
            continue
        diff = ustore.compute_month_diff(official, entries)
        out.append({
            "doctor": doctor,
            "month": mk,
            "added": diff["added_count"],
            "removed": diff["removed_count"],
            "updated_at": month.get("updated_at") or "",
        })
    return out
