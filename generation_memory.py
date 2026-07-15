# -*- coding: utf-8 -*-
"""Storage helpers for generated schedule memory.

The memory keeps multiple generated versions. Only active versions are used, and
only assignments before the period being generated count as prior usage.
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any, Optional

import openpyxl

import shift_history


SCHEMA_VERSION = 1
MEMORY_PATH_DEFAULT = "data/generated_shift_memory.json"
OPERATIONAL_EXCLUDED_DOCTORS = {"Recupero"}


def empty_memory() -> dict:
    return {"schema_version": SCHEMA_VERSION, "versions": []}


def _date_iso(value: dt.date | str) -> str:
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return dt.date.fromisoformat(str(value)[:10]).isoformat()


def _parse_date(value: Any) -> Optional[dt.date]:
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except Exception:
        return None


def normalize_memory(memory: dict | None) -> dict:
    if not isinstance(memory, dict):
        return empty_memory()
    versions = memory.get("versions")
    if not isinstance(versions, list):
        versions = []
    out = empty_memory()
    for version in versions:
        if not isinstance(version, dict):
            continue
        assignments = version.get("assignments")
        if not isinstance(assignments, dict):
            assignments = {}
        out["versions"].append({
            "id": str(version.get("id") or ""),
            "label": str(version.get("label") or ""),
            "start_date": str(version.get("start_date") or "")[:10],
            "end_date": str(version.get("end_date") or "")[:10],
            "created_at": str(version.get("created_at") or ""),
            "active": bool(version.get("active", True)),
            "assignments": assignments,
            "source": str(version.get("source") or "streamlit"),
        })
    return out


def append_version(
    memory: dict | None,
    *,
    version_id: str,
    label: str,
    start_date: dt.date | str,
    end_date: dt.date | str,
    assignments: dict[str, dict[str, list[str]]],
    active: bool = True,
    created_at: Optional[str] = None,
    source: str = "streamlit",
) -> dict:
    out = normalize_memory(memory)
    out["versions"].append({
        "id": str(version_id),
        "label": str(label or version_id),
        "start_date": _date_iso(start_date),
        "end_date": _date_iso(end_date),
        "created_at": created_at or dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
        "active": bool(active),
        "assignments": assignments or {},
        "source": source,
    })
    return out


def memory_to_text(memory: dict) -> str:
    return json.dumps(normalize_memory(memory), ensure_ascii=False, indent=2, sort_keys=False)


def memory_from_text(text: str) -> dict:
    if not str(text or "").strip():
        return empty_memory()
    return normalize_memory(json.loads(text))


def set_version_active(memory: dict | None, active_by_id: dict[str, bool]) -> dict:
    out = normalize_memory(memory)
    for version in out["versions"]:
        vid = str(version.get("id") or "")
        if vid in active_by_id:
            version["active"] = bool(active_by_id[vid])
    return out


def _target_month_cutoffs(start_date: dt.date, end_date: dt.date) -> dict[str, dt.date]:
    cutoffs: dict[str, dt.date] = {}
    cur = start_date
    while cur <= end_date:
        mk = f"{cur.year:04d}-{cur.month:02d}"
        cutoffs.setdefault(mk, cur)
        cur += dt.timedelta(days=1)
    return cutoffs


def _effective_active_assignments(
    memory: dict | None,
    selected_version_ids: Optional[set[str]] = None,
) -> list[tuple[dt.date, dict, str]]:
    """Return effective assignments, with the latest selected/active version winning per date."""
    mem = normalize_memory(memory)
    selected = {str(v) for v in selected_version_ids} if selected_version_ids is not None else None
    by_date: dict[str, tuple[dict, str]] = {}
    for version in mem.get("versions", []):
        version_id = str(version.get("id") or "")
        if selected is None:
            if not version.get("active", True):
                continue
        else:
            if version_id not in selected:
                continue
        assignments = version.get("assignments") or {}
        if not isinstance(assignments, dict):
            continue
        for date_s, by_col in assignments.items():
            day = _parse_date(date_s)
            if day is None or not isinstance(by_col, dict):
                continue
            by_date[day.isoformat()] = (by_col, version_id)

    out: list[tuple[dt.date, dict, str]] = []
    for date_s in sorted(by_date):
        by_col, version_id = by_date[date_s]
        out.append((dt.date.fromisoformat(date_s), by_col, version_id))
    return out


def build_solver_prior_usage(
    memory: dict | None,
    start_date: dt.date,
    end_date: dt.date,
    *,
    finalized_months: Optional[set[str]] = None,
    selected_version_ids: Optional[set[str]] = None,
) -> dict:
    """Build compact prior usage for the solver.

    Only active versions are considered. For each target month, only assignments
    with date < first generated date in that month are counted. This prevents a
    saved version of the same period from counting against its own regeneration.
    """
    cutoffs = _target_month_cutoffs(start_date, end_date)
    finalized = set(finalized_months or set())
    counts: dict[str, dict[str, dict[str, int]]] = {}
    night_dates: dict[str, dict[str, list[str]]] = {}
    versions_used: list[str] = []

    for day, by_col, version_id in _effective_active_assignments(memory, selected_version_ids=selected_version_ids):
        used_version = False
        mk = f"{day.year:04d}-{day.month:02d}"
        if mk in finalized:
            continue
        cutoff = cutoffs.get(mk)
        if cutoff is None:
            for target_mk, target_cutoff in cutoffs.items():
                delta = (target_cutoff - day).days
                if 0 < delta <= 31:
                    docs = by_col.get("J") or by_col.get("j") or []
                    if isinstance(docs, list):
                        for doc in docs:
                            doc_s = str(doc).strip()
                            if not doc_s or doc_s in OPERATIONAL_EXCLUDED_DOCTORS:
                                continue
                            night_dates.setdefault(target_mk, {}).setdefault(doc_s, [])
                            ds = day.isoformat()
                            if ds not in night_dates[target_mk][doc_s]:
                                night_dates[target_mk][doc_s].append(ds)
                            used_version = True
            if used_version and version_id and version_id not in versions_used:
                versions_used.append(version_id)
            continue
        if day >= cutoff:
            continue
        for col, docs in by_col.items():
            col_s = str(col).strip().upper()
            if not col_s or not isinstance(docs, list):
                continue
            for doc in docs:
                doc_s = str(doc).strip()
                if not doc_s or doc_s in OPERATIONAL_EXCLUDED_DOCTORS:
                    continue
                counts.setdefault(mk, {}).setdefault(doc_s, {})
                counts[mk][doc_s][col_s] = counts[mk][doc_s].get(col_s, 0) + 1
                if col_s == "J":
                    night_dates.setdefault(mk, {}).setdefault(doc_s, [])
                    ds = day.isoformat()
                    if ds not in night_dates[mk][doc_s]:
                        night_dates[mk][doc_s].append(ds)
                used_version = True
        if used_version and version_id and version_id not in versions_used:
            versions_used.append(version_id)

    for by_doc in night_dates.values():
        for dates in by_doc.values():
            dates.sort()
    return {
        "counts": counts,
        "night_dates_by_doc": night_dates,
        "versions_used": versions_used,
    }


def memory_to_shift_history(
    memory: dict | None,
    *,
    start_date: Optional[dt.date] = None,
    end_date: Optional[dt.date] = None,
    finalized_months: Optional[set[str]] = None,
    valid_doctors: Optional[set[str]] = None,
    selected_version_ids: Optional[set[str]] = None,
) -> dict[str, dict]:
    """Convert active generated versions into shift-history month stats.

    This returns the same month-stat shape produced by
    `shift_history.compute_doctor_stats()`, so it can be merged with the
    definitive uploaded history before aggregation.

    If `start_date` is provided, only assignments before that date are included.
    That keeps regeneration of the same period from counting its own previous
    version, while allowing a generated first week to influence the rest of the
    month.
    """
    finalized = set(finalized_months or set())
    by_month: dict[str, dict] = {}
    version_ids_by_month: dict[str, set[str]] = {}

    for day, by_col, version_id in _effective_active_assignments(memory, selected_version_ids=selected_version_ids):
        if start_date is not None and day >= start_date:
            continue
        if end_date is not None and start_date is None and day > end_date:
            continue
        mk = f"{day.year:04d}-{day.month:02d}"
        if mk in finalized:
            continue
        parsed_month = by_month.setdefault(
            mk,
            {
                "year": day.year,
                "month": day.month,
                "month_label": mk,
                "days": {},
            },
        )
        day_key = day.isoformat()
        day_rec = parsed_month["days"].setdefault(
            day_key,
            {
                "date": day_key,
                "dow": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][day.weekday()],
                "is_holiday": shift_history._is_holiday(dt.datetime(day.year, day.month, day.day)),
                "assignments": {},
            },
        )
        for col, docs in by_col.items():
            col_s = str(col).strip().upper()
            if not col_s or not isinstance(docs, list):
                continue
            clean_docs = [
                str(d).strip()
                for d in docs
                if str(d).strip() and str(d).strip() not in OPERATIONAL_EXCLUDED_DOCTORS
            ]
            if not clean_docs:
                continue
            day_rec["assignments"].setdefault(col_s, [])
            day_rec["assignments"][col_s].extend(clean_docs)
        if version_id and any(day_rec["assignments"].values()):
            version_ids_by_month.setdefault(mk, set()).add(version_id)

    out: dict[str, dict] = {}
    for mk, parsed in sorted(by_month.items()):
        days_map = parsed.get("days") or {}
        parsed_for_stats = {
            "year": parsed.get("year"),
            "month": parsed.get("month"),
            "month_label": mk,
            "days": [days_map[k] for k in sorted(days_map)],
        }
        stats = shift_history.compute_doctor_stats(parsed_for_stats, valid_doctors=valid_doctors)
        if not any(k != "_meta" for k in stats):
            continue
        last_night: list[str] = []
        if parsed_for_stats["days"]:
            last_night = parsed_for_stats["days"][-1].get("assignments", {}).get("J", [])
            if valid_doctors:
                last_night = [d for d in last_night if d in valid_doctors]
        stats["_meta"] = {
            "source": "generation_memory",
            "partial": True,
            "generation_versions_used": sorted(version_ids_by_month.get(mk, set())),
            "last_day_night_doctors": last_night,
        }
        out[mk] = stats
    return out


def _parse_generated_cell(value: Any) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    parts: list[str] = []
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        if "/" in line and not line[:5].replace("/", "").isdigit():
            parts.extend(p.strip() for p in line.split("/") if p.strip())
        else:
            parts.append(line)

    out: list[str] = []
    for part in parts:
        base = part.split("(", 1)[0].strip()
        if base.casefold() == "recupero":
            out.append("Recupero")
            continue
        name = shift_history._normalize_name(part)
        if name:
            out.append(name)
    return out


def parse_generated_xlsx_assignments(xlsx_path: str | Path, sheet_name: Optional[str] = None) -> dict[str, dict[str, list[str]]]:
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    try:
        if sheet_name:
            ws = wb[sheet_name]
        else:
            for name in wb.sheetnames:
                if name.lower() != "riepilogo":
                    ws = wb[name]
                    break
            else:
                ws = wb[wb.sheetnames[0]]

        col_map = shift_history._map_columns_from_header(ws)
        out: dict[str, dict[str, list[str]]] = {}
        for row_idx in range(2, ws.max_row + 1):
            date_val = ws.cell(row_idx, 1).value
            if date_val is None:
                break
            if isinstance(date_val, dt.datetime):
                day = date_val.date()
            elif isinstance(date_val, dt.date):
                day = date_val
            else:
                continue
            cleaned: dict[str, list[str]] = {}
            for col_idx, tag in col_map.items():
                names = _parse_generated_cell(ws.cell(row_idx, col_idx).value)
                if names:
                    cleaned[str(tag).strip().upper()] = names
            if cleaned:
                out[day.isoformat()] = cleaned
        return out
    finally:
        wb.close()


def load_memory_from_github(
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    path: str = MEMORY_PATH_DEFAULT,
) -> tuple[dict, Optional[str]]:
    from github_utils import get_file

    gf = get_file(owner, repo, path, token, branch)
    if gf is None:
        return empty_memory(), None
    return memory_from_text(gf.text), gf.sha


def save_memory_to_github(
    memory: dict,
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    sha: Optional[str] = None,
    path: str = MEMORY_PATH_DEFAULT,
) -> dict:
    from github_utils import put_file

    return put_file(
        owner=owner,
        repo=repo,
        path=path,
        token=token,
        branch=branch,
        sha=sha,
        message="Update generated shift memory",
        text=memory_to_text(memory),
    )
