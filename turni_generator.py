#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Turni Autogenerator – UTIC/Cardiologia
- Legge un template Excel (openpyxl)
- Legge regole YAML
- Legge indisponibilità mensili (Excel/CSV) con colonne: Medico, Data, Fascia (Mattina/Pomeriggio/Notte/Diurno/Tutto il giorno)
- Compila le colonne operative e salva un nuovo file Excel
Solver principale: OR-Tools CP-SAT (pip install ortools)
Fallback: greedy + report conflitti (meno robusto)
Autore: prototype by ChatGPT (GPT-5.2 Thinking)
"""



from __future__ import annotations
__version__ = "2026-01-24-freecols-highlight-1"

import argparse
import dataclasses
import calendar
import datetime as dt
import math
import re
import sys
from collections import defaultdict, Counter
from copy import deepcopy, copy as _copy
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple, Iterable
import pandas as pd
import yaml
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill
# -------------------------
# Utilities
# -------------------------
DOW_MAP = {
    0: "Mon",
    1: "Tue",
    2: "Wed",
    3: "Thu",
    4: "Fri",
    5: "Sat",
    6: "Sun",
}

# Optional style template used to make auto-generated monthly templates look
# like the official hospital model (header, column widths, weekend/holiday
# shading, etc.).
#
# Place a file named `Style_Template.xlsx` in the repo root (same folder as
# this script) to enable styling.
STYLE_TEMPLATE_FILENAME = "Style_Template.xlsx"
STYLE_TEMPLATE_FALLBACK = "Modello_Febbraio_2026.xlsx"  # backward-compat if you keep the old name


def _find_style_template() -> Optional[Path]:
    base = Path(__file__).resolve().parent
    p1 = base / STYLE_TEMPLATE_FILENAME
    if p1.exists():
        return p1
    p2 = base / STYLE_TEMPLATE_FALLBACK
    if p2.exists():
        return p2
    return None


def _load_style_ws() -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
    """Load the first worksheet from the style template, if present."""
    p = _find_style_template()
    if not p:
        return None
    try:
        wb = openpyxl.load_workbook(p)
        return wb[wb.sheetnames[0]]
    except Exception:
        return None


def _easter_date_gregorian(year: int) -> dt.date:
    """Compute Easter Sunday (Gregorian calendar) using the Meeus/Jones/Butcher algorithm."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return dt.date(year, month, day)


def italy_public_holidays(year: int) -> Set[dt.date]:
    """Italian national public holidays (incl. Easter Monday)."""
    easter = _easter_date_gregorian(year)
    easter_monday = easter + dt.timedelta(days=1)
    fixed = {
        dt.date(year, 1, 1),   # Capodanno
        dt.date(year, 1, 6),   # Epifania
        dt.date(year, 4, 25),  # Liberazione
        dt.date(year, 5, 1),   # Lavoro
        dt.date(year, 6, 2),   # Repubblica
        dt.date(year, 8, 15),  # Ferragosto
        dt.date(year, 11, 1),  # Ognissanti
        dt.date(year, 12, 8),  # Immacolata
        dt.date(year, 12, 25), # Natale
        dt.date(year, 12, 26), # Santo Stefano
    }
    return fixed | {easter_monday}


def _is_grey_solid(cell) -> bool:
    """Return True if cell has the same grey solid fill used by the model."""
    try:
        fill = cell.fill
        if not fill or fill.patternType != "solid":
            return False
        fg = getattr(fill.fgColor, "rgb", None)
        return fg == "FFC0C0C0"
    except Exception:
        return False


def _copy_cell_style(src, dst) -> None:
    """Copy style elements from src cell to dst cell (without copying value)."""
    try:
        # IMPORTANT: openpyxl style objects do not play well with deepcopy
        # (it can recurse). Use shallow copies or direct assignment.
        dst.font = _copy(src.font)
        dst.fill = _copy(src.fill)
        dst.border = _copy(src.border)
        dst.alignment = _copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = _copy(src.protection)
    except Exception:
        # best-effort: ignore style copy failures
        pass


def _apply_model_style_to_template(ws, cfg: dict, year: int, month: int, last_day: int) -> None:
    """Apply header/column widths and weekend/holiday shading using Style_Template.xlsx if available."""
    style_ws = _load_style_ws()
    if style_ws is None:
        return

    # Identify representative rows in the style worksheet
    # - sunday_row: first row where column B contains 'domenica' (Italian)
    # - weekday_row: first data row that is not sunday_row
    sunday_row = None
    weekday_row = None
    for r in range(2, min(style_ws.max_row, 20) + 1):
        b = style_ws.cell(r, 2).value
        if isinstance(b, str) and b.strip().lower().startswith("domen"):
            sunday_row = r
            break
    # fallback: assume row 2 is sunday in the provided model
    if sunday_row is None:
        sunday_row = 2
    # weekday row
    for r in range(2, min(style_ws.max_row, 20) + 1):
        if r == sunday_row:
            continue
        a = style_ws.cell(r, 1).value
        if isinstance(a, (dt.date, dt.datetime)):
            weekday_row = r
            break
    if weekday_row is None:
        weekday_row = min(3, style_ws.max_row)

    # Column widths + header styles/labels
    style_max_col = style_ws.max_column

    # The auto-template can declare columns beyond the style model (e.g. adding
    # extra "Medici liberi" columns AF/AG). In that case we still want the new
    # columns to look consistent: we extend styles by cloning from the last
    # available style column.
    needed_max_col = style_max_col
    try:
        cols_map = cfg.get("columns") or {}
        keep_empty = cfg.get("keep_empty_columns") or []
        cand_letters = []
        if isinstance(cols_map, dict):
            cand_letters.extend(list(cols_map.keys()))
        if isinstance(keep_empty, list):
            cand_letters.extend(list(keep_empty))
        for _cl in cand_letters:
            try:
                idx = column_index_from_string(str(_cl).strip().upper())
                if idx > needed_max_col:
                    needed_max_col = idx
            except Exception:
                pass
    except Exception:
        pass

    max_col = max(style_max_col, needed_max_col)
    # Ensure we have at least up to max_col in row 1
    _ = ws.cell(row=1, column=max_col)

    # Row heights
    if style_ws.row_dimensions[1].height:
        ws.row_dimensions[1].height = style_ws.row_dimensions[1].height
    if style_ws.row_dimensions[weekday_row].height:
        data_h = style_ws.row_dimensions[weekday_row].height
        for r in range(2, last_day + 2):
            ws.row_dimensions[r].height = data_h

    # Copy column widths and header style/value
    # For columns beyond the style model, clone from the last style column.
    ref_c = style_max_col
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        src_idx = c if c <= style_max_col else ref_c
        src_letter = get_column_letter(src_idx)
        w = style_ws.column_dimensions[src_letter].width
        if w:
            ws.column_dimensions[letter].width = w

        src_h = style_ws.cell(1, src_idx)
        dst_h = ws.cell(1, c)
        _copy_cell_style(src_h, dst_h)
        # Fill missing header labels from model (important for empty spacer columns)
        if (dst_h.value is None or str(dst_h.value).strip() == "") and (src_h.value is not None):
            dst_h.value = src_h.value

    # Determine which columns are shaded in the model on Sundays/holidays.
    # Extra columns (beyond the model) inherit the last-model-column shading.
    grey_cols = {c for c in range(1, style_max_col + 1) if _is_grey_solid(style_ws.cell(sunday_row, c))}
    try:
        if max_col > style_max_col and _is_grey_solid(style_ws.cell(sunday_row, ref_c)):
            grey_cols |= set(range(style_max_col + 1, max_col + 1))
    except Exception:
        pass

    # Holiday set for styling. The template may cover a custom range across
    # multiple months/years, so derive involved years from the actual rows.
    extra = set()
    for x in cfg.get("festivi_extra", []) or []:
        try:
            extra.add(parse_date(x))
        except Exception:
            pass
    row_dates = []
    for r in range(2, last_day + 2):
        d = ws.cell(r, 1).value
        if isinstance(d, dt.datetime):
            d = d.date()
        if isinstance(d, dt.date):
            row_dates.append(d)
    years = {d.year for d in row_dates} or {int(year)}
    holidays = set()
    for y in years:
        holidays |= italy_public_holidays(int(y))
    holidays |= extra

    # Apply per-cell styles for the month (lightweight: <= 31 rows * ~31 cols)
    for r in range(2, last_day + 2):
        d = ws.cell(r, 1).value
        if isinstance(d, dt.datetime):
            d = d.date()
        if not isinstance(d, dt.date):
            continue
        is_holiday = (d.weekday() == 6) or (d in holidays)
        for c in range(1, max_col + 1):
            dst = ws.cell(r, c)
            src_idx = c if c <= style_max_col else ref_c
            # Choose style row based on holiday shading columns
            if is_holiday and (c in grey_cols):
                src = style_ws.cell(sunday_row, src_idx)
            else:
                src = style_ws.cell(weekday_row, src_idx)
            _copy_cell_style(src, dst)

SHIFT_NORMALIZE = {
    "mattina": "Mattina",
    "pom": "Pomeriggio",
    "pomeriggio": "Pomeriggio",
    "notte": "Notte",
}
def parse_date(x) -> dt.date:
    """Parse a date from Excel/str/datetime."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        raise ValueError("Empty date")
    if isinstance(x, dt.date) and not isinstance(x, dt.datetime):
        return x
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, (int, float)):
        # excel serial date: pandas handles it better; fallback not used here
        raise ValueError(f"Numeric date not supported directly: {x}")
    s = str(x).strip()
    # Accept dd/mm/yyyy or yyyy-mm-dd
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unrecognized date format: {x}")
def norm_name(x: str) -> str:
    return re.sub(r"\s+", " ", str(x).strip())
def norm_shift(x: str) -> str:
    s = str(x).strip().lower()
    if s in SHIFT_NORMALIZE:
        return SHIFT_NORMALIZE[s]
    # allow first letter
    if s.startswith("m"):
        return "Mattina"
    if s.startswith("p"):
        return "Pomeriggio"
    if s.startswith("n"):
        return "Notte"
    raise ValueError(f"Unrecognized shift: {x}")
def shifts_from_fascia(x: str) -> Set[str]:
    """Map unavailability 'Fascia' value to one or more internal shifts.
    Accepted:
      - Mattina / Pomeriggio / Notte  -> that single shift
      - Diurno (or Giorno)            -> {'Mattina','Pomeriggio'}
      - Tutto il giorno / All day     -> {'Any'} (treated as full-day)
      - Ferie / Vacanza / Holiday / Leave -> {'Any'} (treated as full-day)
    """
    s = str(x).strip().lower()
    # Full-day first (because it contains 'giorno')
    if any(k in s for k in ["tutto", "intera", "completa", "allday", "all day", "full day", "24h", "24 h", "ferie", "vacan", "holiday", "leave"]):
        return {"Any"}
    # Daytime (morning + afternoon, but still allows night)
    if s in ["diurno", "giorno", "daytime", "day"] or s.startswith("diurn"):
        return {"Mattina", "Pomeriggio"}
    # Default single-shift
    return {norm_shift(x)}
def dayspec_contains(dow: str, spec) -> bool:
    """
    spec can be:
    - string: 'Mon-Sat', 'Mon-Fri', 'Wed'
    - list of strings: ['Mon','Tue']
    - None: means all days
    """
    if spec is None:
        return True
    if isinstance(spec, str):
        spec = spec.strip()
        if "-" in spec:
            a, b = spec.split("-", 1)
            a, b = a.strip(), b.strip()
            order = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
            ia, ib = order.index(a), order.index(b)
            if ia <= ib:
                return order.index(dow) >= ia and order.index(dow) <= ib
            # wrap (rare)
            return order.index(dow) >= ia or order.index(dow) <= ib
        return dow == spec
    if isinstance(spec, list):
        return dow in spec
    raise TypeError(f"Invalid days spec: {spec}")
# -------------------------
# Data structures
# -------------------------
@dataclasses.dataclass(frozen=True)
class DayRow:
    date: dt.date
    dow: str  # 'Mon'...'Sun'
    row_idx: int
@dataclasses.dataclass
class Slot:
    """One assignable decision for a day."""
    day: DayRow
    slot_id: str                 # unique id
    columns: List[str]           # Excel column letters to fill with same doctor
    allowed: List[str]           # allowed doctor names (domain)
    required: bool = True        # if False, can be blank
    blank_penalty: int = 0      # penalty if left blank (only if required=False)
    shift: str = "Mattina"       # Mattina / Pomeriggio / Notte / Any
    rule_tag: str = ""           # for reporting
    empty_domain: bool = False    # True if allowed domain becomes empty after applying unavailability

    force_same_doctor: bool = False  # if True, this slot must use same doctor as its paired slot (e.g., D=F fallback)
    emergency_doctors: Optional[List[str]] = None  # medici di emergenza (non nel pool primario) — usati con penalità
# -------------------------
# Reperibilità (C) – final assignment layer
# -------------------------

def assign_reperibilita_C(cfg: dict, days: List[DayRow], slots: List[Slot],
                         assignment: Dict[str, Optional[str]]) -> Tuple[Dict[str, Optional[str]], Dict]:
    """Assign/Reassign Reperibilità (column C) as a final layer (STRICT).

    HARD rules (no automatic relaxation):
      - Exactly one doctor per day on column C (if C_reperibilita is configured).
      - C CAN overlap with other tasks on weekdays.
      - C is NOT allowed if:
          1) doctor has Night (J) on the same day
          2) doctor has Night (J) the previous day
          3) on Sundays/holidays, doctor is already working any other task that same day
      - Per doctor per month: min_per_doctor and max_per_doctor (typ. 2..3)
      - Minimum spacing between two C for the same doctor: spacing_min_days (typ. 3)

    If constraints are infeasible, this function raises ValueError with an actionable message.
    """
    rules = cfg.get("rules", {}) or {}
    rC = rules.get("C_reperibilita", {}) if isinstance(rules.get("C_reperibilita", {}), dict) else {}
    if not rC:
        return assignment, {"C_reperibilita_diag": {"status": "SKIPPED", "reason": "C_reperibilita not configured"}}

    constraints = set(rC.get("constraints") or [])
    excluded = {norm_name(x) for x in (rC.get("excluded") or [])}
    spacing_min = int(rC.get("spacing_min_days", 0) or 0)

    # Festivi pool: doctors preferred for C on festivo days (excluding those already
    # excluded from C, e.g. De Gregorio, Manganaro).
    # Doctors doing Notte that day are already blocked by the not_night_same_day constraint.
    rFest = rules.get("Festivi") if isinstance(rules.get("Festivi"), dict) else {}
    festivi_pool_norm = {norm_name(d) for d in (rFest.get("pool") or [])} - excluded
    min_per = int(rC.get("min_per_doctor", rC.get("target_per_doctor", 0) or 0) or 0)
    max_per = int(rC.get("max_per_doctor", 0) or 0)
    target = int(rC.get("target_per_doctor", 0) or 0)
    if days:
        try:
            import calendar as _calendar_c
            _dates_c = sorted(d.date for d in days)
            _first_c = _dates_c[0]
            _last_c = _dates_c[-1]
            _full_month_days_c = _calendar_c.monthrange(_first_c.year, _first_c.month)[1]
            _partial_c = (
                _first_c.day != 1
                or _last_c.day != _full_month_days_c
                or len(_dates_c) != _full_month_days_c
            )
            if _partial_c and _full_month_days_c > 0:
                _frac_c = len(_dates_c) / _full_month_days_c
                target = int(math.floor(target * _frac_c + 0.5))
                min_per = int(math.floor(min_per * _frac_c + 0.5))
                max_per = max(1, int(math.ceil(max_per * _frac_c))) if max_per > 0 else 0
        except Exception:
            pass

    night_col = "J"  # fixed: Notte is column J

    # Map slots by day
    slots_by_day: Dict[dt.date, List[Slot]] = defaultdict(list)
    for s in slots:
        slots_by_day[s.day.date].append(s)

    # Identify C slots
    cslot_by_date: Dict[dt.date, Slot] = {}
    c_dates: List[dt.date] = []
    for d in days:
        cslot = next((s for s in slots_by_day.get(d.date, []) if s.columns == ["C"]), None)
        if cslot is not None:
            cslot_by_date[d.date] = cslot
            c_dates.append(d.date)

    if not c_dates:
        return assignment, {"C_reperibilita_diag": {"status": "SKIPPED", "reason": "No C slots in template"}}

    # Helper: does doctor work in any non-C slot the same day?
    def doctor_works_same_day(date_: dt.date, doc: str) -> bool:
        for s in slots_by_day.get(date_, []):
            if s.columns == ["C"]:
                continue
            if assignment.get(s.slot_id) == doc:
                return True
        return False

    # Day lookup
    dayrow_by_date = {d.date: d for d in days}

    # Allowed candidates per day
    allowed_norm_by_date: Dict[dt.date, List[str]] = {}
    pool_set: Set[str] = set()
    for d in c_dates:
        cand = []
        for raw in (cslot_by_date[d].allowed or []):
            dn = norm_name(raw)
            if not dn or dn in excluded or dn == "Recupero":
                continue
            cand.append(dn)
        # unique, stable
        seen=set(); cand2=[]
        for x in cand:
            if x not in seen:
                seen.add(x); cand2.append(x)
        allowed_norm_by_date[d] = cand2
        pool_set |= set(cand2)

    pool = sorted(pool_set, key=lambda s: s.lower())
    if not pool:
        raise ValueError("C_reperibilita: pool vuoto (tutti esclusi o non presenti nei pools YAML).")

    # Candidate filter applying hard constraints
    def ok_candidate(date_: dt.date, doc: str) -> bool:
        if doc not in set(allowed_norm_by_date.get(date_, [])):
            return False
        if "not_night_same_day" in constraints:
            if assignment.get(f"{date_}-{night_col}") == doc:
                return False
        if "not_night_prev_day" in constraints:
            prev = date_ - dt.timedelta(days=1)
            if assignment.get(f"{prev}-{night_col}") == doc:
                return False
        if "not_night_prev_2_days" in constraints:
            for delta in [1, 2]:
                prev = date_ - dt.timedelta(days=delta)
                if assignment.get(f"{prev}-{night_col}") == doc:
                    return False
        if "not_night_next_2_days" in constraints:
            for delta in [1, 2]:
                nxt = date_ + dt.timedelta(days=delta)
                if assignment.get(f"{nxt}-{night_col}") == doc:
                    return False
        if "not_working_same_day_on_sundays_and_holidays" in constraints:
            drow = dayrow_by_date.get(date_)
            if drow is not None and is_festivo(drow, cfg):
                if doctor_works_same_day(date_, doc):
                    return False
        return True

    # ── Funzione candidati con proximity J configurabile ──────────────────────
    def build_candidates_with_prox(night_prox: int) -> Dict[dt.date, List[str]]:
        """Costruisce candidates_by_date con finestra J±night_prox (0 = solo same-day)."""
        result: Dict[dt.date, List[str]] = {}
        for d in c_dates:
            cands = []
            for doc in pool:
                if doc not in set(allowed_norm_by_date.get(d, [])):
                    continue
                # not_night_same_day sempre attivo
                if assignment.get(f"{d}-{night_col}") == doc:
                    continue
                # Proximity J rilassabile
                blocked = False
                for delta in range(1, night_prox + 1):
                    if assignment.get(f"{(d - dt.timedelta(days=delta))}-{night_col}") == doc:
                        blocked = True; break
                    if assignment.get(f"{(d + dt.timedelta(days=delta))}-{night_col}") == doc:
                        blocked = True; break
                if blocked:
                    continue
                # Festivi: non assegnato nello stesso giorno
                if "not_working_same_day_on_sundays_and_holidays" in constraints:
                    drow = dayrow_by_date.get(d)
                    if drow is not None and is_festivo(drow, cfg):
                        if doctor_works_same_day(d, doc):
                            continue
                cands.append(doc)
            result[d] = cands
        return result

    # Costruisci candidati con proximity progressivamente rilassata
    _c_relaxation_warnings: List[str] = []
    candidates_by_date: Dict[dt.date, List[str]] = {}
    for _night_prox in [2, 1, 0]:
        candidates_by_date = build_candidates_with_prox(_night_prox)
        _empty_days = [d for d in c_dates if not candidates_by_date[d]]
        if not _empty_days:
            if _night_prox < 2:
                _c_relaxation_warnings.append(
                    f"C reperibilità: vincolo J proximity rilassato a ±{_night_prox} giorni "
                    f"per trovare candidati."
                )
            break
        if _night_prox == 0:
            # Ultimo resort: ignora del tutto la disponibilità per i giorni problematici
            for d in _empty_days:
                fallback = list(allowed_norm_by_date.get(d, [])) or pool[:]
                candidates_by_date[d] = fallback
                _c_relaxation_warnings.append(
                    f"C reperibilità: {d} senza candidati strict → usato pool completo ({len(fallback)} dottori)."
                )

    total_days = len(c_dates)
    n_docs = len(pool)

    # Feasibility checks for min/max (con auto-aumento max_per se necessario)
    if max_per <= 0:
        raise ValueError("C_reperibilita: max_per_doctor deve essere > 0.")
    if min_per < 0:
        min_per = 0

    if total_days > n_docs * max_per:
        # Auto-rilassa max_per invece di fallire
        max_per = (total_days + n_docs - 1) // n_docs  # ceil(total/n)
        _c_relaxation_warnings.append(
            f"C reperibilità: max_per_doctor aumentato a {max_per} per coprire {total_days} giorni con {n_docs} medici."
        )
    if min_per > 0 and total_days < n_docs * min_per:
        min_per = 0
        _c_relaxation_warnings.append(
            f"C reperibilità: min_per_doctor azzerato — troppi medici nel pool per il mese."
        )

    # Build desired counts: start at min_per, distribute remaining +1 up to max_per
    desired = {doc: (min_per if min_per > 0 else 0) for doc in pool}
    cur = sum(desired.values())
    remaining = total_days - cur
    order_docs = pool[:]
    while remaining > 0:
        progress = False
        for doc in order_docs:
            if desired[doc] < max_per and remaining > 0:
                desired[doc] += 1
                remaining -= 1
                progress = True
        if not progress:
            break
    if sum(desired.values()) != total_days:
        # Distribuisci il residuo ignorando max_per
        for doc in order_docs:
            if remaining <= 0:
                break
            desired[doc] += 1
            remaining -= 1
    # Il target/desiderato serve solo a ordinare le preferenze: non può diventare
    # un limite hard. Con molte ferie il pattern ideale 2/3 per medico può essere
    # impossibile anche quando max_per_doctor ha capacità sufficiente.
    hard_cap_by_doc = {doc: max_per for doc in pool}

    # Backtracking DFS con retry su spacing rilassato
    assigned: Dict[dt.date, str] = {}

    def _run_dfs_attempt(sp: int) -> bool:
        _cnt = {doc: 0 for doc in pool}
        _dates_by_doc: Dict[str, List[dt.date]] = {doc: [] for doc in pool}
        _c_dates_sorted = sorted(c_dates, key=lambda d: (len(candidates_by_date[d]), d))

        def _spacing_ok(doc: str, d: dt.date) -> bool:
            if sp > 1:
                for prev in _dates_by_doc[doc]:
                    if abs((d - prev).days) < sp:
                        return False
            return True

        def _pick(d: dt.date) -> List[str]:
            cands = candidates_by_date[d]
            drow = dayrow_by_date.get(d)
            day_is_festivo = drow is not None and is_festivo(drow, cfg)
            def key(doc):
                festivo_pref = 1 if (day_is_festivo and doc in festivi_pool_norm) else 0
                return (festivo_pref, desired[doc] - _cnt[doc], -_cnt[doc], doc.lower())
            return sorted(cands, key=key, reverse=True)

        def _dfs(i: int) -> bool:
            if i == len(_c_dates_sorted):
                return True
            d = _c_dates_sorted[i]
            for doc in _pick(d):
                if _cnt[doc] >= hard_cap_by_doc[doc]:
                    continue
                if not _spacing_ok(doc, d):
                    continue
                assigned[d] = doc
                _cnt[doc] += 1
                _dates_by_doc[doc].append(d)
                if _dfs(i + 1):
                    return True
                _dates_by_doc[doc].pop()
                _cnt[doc] -= 1
                assigned.pop(d, None)
            return False

        assigned.clear()
        return _dfs(0)

    solved = False
    _spacing_tried = spacing_min
    _cap_tried = max_per
    cap_values = list(range(max_per, total_days + 1))
    for _cap in cap_values:
        hard_cap_by_doc = {doc: _cap for doc in pool}
        for _sp in ([spacing_min] + list(range(spacing_min - 1, -1, -1))):
            solved = _run_dfs_attempt(_sp)
            if solved:
                _spacing_tried = _sp
                _cap_tried = _cap
                if _sp < spacing_min:
                    _c_relaxation_warnings.append(
                        f"C reperibilità: spacing rilassato a {_sp} giorni (originale {spacing_min})."
                    )
                if _cap > max_per:
                    _c_relaxation_warnings.append(
                        f"C reperibilità: max_per_doctor rilassato a {_cap} per compatibilità con ferie/vincoli."
                    )
                break
        if solved:
            break

    if not solved:
        raise ValueError(
            "C_reperibilita: impossibile assegnare la reperibilità anche con vincoli rilassati. "
            f"Pool: {pool}, pool_size={n_docs}, giorni={total_days}, "
            f"max_per_doctor provato fino a {total_days}. "
            "Verifica giorni senza candidati o esclusioni del pool C."
        )

    # Write back into assignment
    for d in c_dates:
        assignment[cslot_by_date[d].slot_id] = assigned.get(d)

    # Diagnostics (compute counts from assigned dict)
    _assigned_cnt: Dict[str, int] = {}
    for d in c_dates:
        doc = assigned.get(d)
        if doc:
            _assigned_cnt[doc] = _assigned_cnt.get(doc, 0) + 1

    status_str = "OK_RELAXED" if _c_relaxation_warnings else "OK_STRICT"
    diag: Dict = {
        "status": status_str,
        "pool_size": n_docs,
        "total_days": total_days,
        "spacing_min_days": spacing_min,
        "effective_spacing_days": _spacing_tried,
        "effective_max_per_doctor": _cap_tried,
    }
    diag["counts"] = {k: v for k, v in sorted(_assigned_cnt.items(), key=lambda kv: (-kv[1], kv[0].lower())) if v}
    if _c_relaxation_warnings:
        diag["relaxation_warnings"] = _c_relaxation_warnings
    # Overlap stats
    overlap_total = 0
    overlap_weekdays = 0
    for d in c_dates:
        doc = assigned.get(d)
        if not doc:
            continue
        works = doctor_works_same_day(d, doc)
        if works:
            overlap_total += 1
            drow = dayrow_by_date.get(d)
            if drow is not None and not is_festivo(drow, cfg):
                overlap_weekdays += 1
    diag["overlap_days_total"] = overlap_total
    diag["overlap_days_weekdays"] = overlap_weekdays

    return assignment, {"C_reperibilita_diag": diag}


# -------------------------
# Load config / template
# -------------------------
def load_rules(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    if not isinstance(cfg, dict):
        raise ValueError("Rules YAML must be a mapping.")
    return cfg


def apply_pool_config(cfg_yaml: dict, pool_cfg: Optional[dict]) -> dict:
    """Sovrascrive pool/quote/flag nel cfg YAML con i valori del pool_config JSON.

    Se pool_cfg è None o vuoto ritorna una deep copy di cfg_yaml invariata.
    La funzione è idempotente: applicata più volte produce lo stesso risultato.
    """
    import copy as _copy

    cfg = _copy.deepcopy(cfg_yaml)
    if not pool_cfg or not pool_cfg.get("doctors"):
        return cfg
    try:
        from pool_config_store import normalize_pool_config as _normalize_pool_config
        pool_cfg = _normalize_pool_config(pool_cfg)
    except Exception:
        pass

    doctors: dict = pool_cfg.get("doctors", {})
    col_settings: dict = pool_cfg.get("column_settings", {})
    rules = cfg.setdefault("rules", {})
    gc = cfg.setdefault("global_constraints", {})

    # 1. active=false → absolute_exclusions
    abs_excl: list = list(cfg.get("absolute_exclusions") or [])
    abs_excl_set = {norm_name(x) for x in abs_excl}
    for doc, dcfg in doctors.items():
        if not dcfg.get("active", True) and norm_name(doc) not in abs_excl_set:
            abs_excl.append(doc)
            abs_excl_set.add(norm_name(doc))
    cfg["absolute_exclusions"] = abs_excl

    # Medici attivi (normalizzati)
    active_docs = [doc for doc, dcfg in doctors.items() if dcfg.get("active", True)]
    active_set = {norm_name(d) for d in active_docs}

    # 2. Mappa colonna → pool key nel YAML (per sostituire i pool)
    _COL_RULE: dict[str, list[tuple[str, str]]] = {
        # (rule_key, pool_field)
        "C":  [("C_reperibilita", None)],         # C si gestisce via excluded, non pool
        # D/F share one solver rule, but D is the primary ward column.
        # A doctor enabled only on F must remain a fallback/support candidate,
        # not become part of the primary Grimaldi/Calabro pair.
        "D":  [("D_F", "allowed")],
        "F":  [],
        "E":  [("E_G", "allowed")],
        "G":  [("E_G", "allowed")],
        "H":  [("H", "pool_mon_fri"), ("H", "distribution_pool")],
        "I":  [("I", "distribution_pool")],
        "J":  [("J", "pool_other")],
        "K":  [("K", "pool")],
        "L":  [("L", "pool_other")],
        "Q":  [("Q", "pool")],
        "R":  [("R", "pool")],
        "S":  [("S", "pool")],
        "T":  [("T", "pool")],
        "U":  [("U", "pool")],
        "V":  [("V", "pool")],
        "W":  [("W", "other_days_pool")],
        "Y":  [("Y", "other_pool")],
        "Z":  [("Z", "pool")],
        "AB": [("AB", "fallback_pool")],
    }

    for col, rule_targets in _COL_RULE.items():
        if col == "C":
            continue  # gestito separatamente al punto 5
        new_pool = [
            doc for doc, dcfg in doctors.items()
            if dcfg.get("active", True) and col in (dcfg.get("columns") or [])
        ]
        for rule_key, pool_field in rule_targets:
            rules.setdefault(rule_key, {})[pool_field] = new_pool

    # 3. Pool festivi diurni (D/E/H/I nei giorni festivi)
    fest_incl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True) and dcfg.get("festivi_diurni", True)
    ]
    fest_excl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True) and not dcfg.get("festivi_diurni", True)
    ]
    if fest_incl:
        rules.setdefault("Festivi", {})["pool"] = fest_incl
    if fest_excl:
        rules.setdefault("Festivi", {})["excluded"] = fest_excl

    # 4. Pool festivi notti (J nei giorni festivi)
    festivi_notti_excl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True)
        and "J" in (dcfg.get("columns") or [])
        and not dcfg.get("festivi_notti", True)
    ]
    cfg["pool_festivi_notti_excluded"] = {norm_name(d) for d in festivi_notti_excl}

    # 4b. Esclusione sabato diurno: vale per turni Mattina/Pomeriggio,
    # non per J notte e non per C reperibilita'.
    saturday_day_excl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True) and dcfg.get("exclude_saturday_day", False)
    ]
    cfg["pool_saturday_day_excluded"] = {norm_name(d) for d in saturday_day_excl}

    # 5. Reperibilità C: sostituisce excluded list
    c_excl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True) and dcfg.get("excluded_from_reperibilita", False)
    ]
    # Aggiungi anche i non-attivi (non assegnabili comunque ma meglio espliciti)
    for doc, dcfg in doctors.items():
        if not dcfg.get("active", True) and doc not in c_excl:
            c_excl.append(doc)
    rules.setdefault("C_reperibilita", {})["excluded"] = c_excl

    # 6. Weekend J excluded — non richiede J in columns (può entrare via monthly_quotas)
    j_weekend_excl = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True)
        and not (dcfg.get("column_overrides") or {}).get("J", {}).get("weekend_nights", True)
    ]
    rules.setdefault("J", {})["weekend_excluded_doctors"] = j_weekend_excl

    # 7. Quote J fixed → J.monthly_quotas (compatibile con solver esistente)
    j_mq: dict = dict(rules.get("J", {}).get("monthly_quotas") or {})
    quota_overrides: dict = {}  # (doc_norm, col) → {type, value}

    for doc, dcfg in doctors.items():
        dn = norm_name(doc)
        overrides: dict = dcfg.get("column_overrides") or {}
        for col, ov in overrides.items():
            if not isinstance(ov, dict):
                continue
            mq = ov.get("monthly_quota")
            qt = ov.get("quota_type", "fixed")
            if mq is None:
                continue
            if col == "J" and qt == "fixed":
                j_mq[doc] = mq  # usa chiave originale per compatibilità
            else:
                quota_overrides[(dn, col)] = {"type": qt, "value": int(mq)}

    rules.setdefault("J", {})["monthly_quotas"] = j_mq
    cfg["pool_quota_overrides"] = quota_overrides

    # 8. counts_as per colonna
    counts_as_map: dict[str, int] = {}
    monthly_targets: dict[str, int] = {}
    for col, cs in col_settings.items():
        if isinstance(cs, dict) and "counts_as" in cs:
            counts_as_map[col] = int(cs["counts_as"])
        if isinstance(cs, dict) and cs.get("monthly_target") is not None:
            try:
                monthly_targets[str(col).strip().upper()] = int(cs.get("monthly_target"))
            except Exception:
                pass
    if counts_as_map:
        cfg["pool_counts_as"] = counts_as_map
    if monthly_targets:
        cfg["pool_monthly_targets"] = monthly_targets

    # 9. service_combinations
    combos = pool_cfg.get("service_combinations")
    if combos is not None:
        cfg["pool_service_combinations"] = combos

    # 10. critical_services
    critical = pool_cfg.get("critical_services")
    if critical is not None:
        cfg["pool_critical_services"] = critical

    # 11. Spacing J: scrive direttamente nei global_constraints già letti dal solver
    j_cs = col_settings.get("J", {})
    if isinstance(j_cs, dict):
        if "spacing_min_days" in j_cs:
            gc["night_spacing_days_min"] = int(j_cs["spacing_min_days"])
        if "spacing_preferred_days" in j_cs:
            gc["night_spacing_days_preferred"] = int(j_cs["spacing_preferred_days"])

    # 11b. service_combinations → mappa alle chiavi relief_valves esistenti
    combos = pool_cfg.get("service_combinations") or []
    relief = gc.setdefault("relief_valves", {})
    for combo in combos:
        cols = tuple(sorted(combo.get("columns") or []))
        mode = combo.get("mode", "fallback")
        if cols == ("K", "T"):
            if mode in ("fallback", "always"):
                relief["enable_kt_share"] = True
            elif mode == "preferred":
                relief["enable_kt_share"] = False  # preferred = soft, solver gestisce a bassa penalità
        # Q+R: già gestito da allow_blank_columns.R — nessuna modifica necessaria

    # 12. university_doctors — aggiorna da pool_config
    gc_uni: dict = dict(gc.get("university_doctors") or {})
    for doc, dcfg in doctors.items():
        uni = dcfg.get("university_doctor")
        if uni and isinstance(uni, dict):
            ratio = float(uni.get("ratio", gc.get("university_ratio", 0.6)))
            night_double = "J" in (dcfg.get("columns") or [])
            gc_uni[doc] = {
                "type": "university",
                "night_counts_double": night_double,
            }
            gc["university_ratio"] = ratio
        elif doc in gc_uni and uni is None:
            # Rimosso da pool_config → rimuovi anche da gc_uni
            del gc_uni[doc]
    gc["university_doctors"] = gc_uni

    # 13. Soft balance per colonne senza balance nel YAML
    # U (Contr.PM) non ha balance:true nel YAML — lo aggiungiamo quando pool_config
    # definisce il pool, così il solver bilancia automaticamente il carico.
    _u_pool = [
        doc for doc, dcfg in doctors.items()
        if dcfg.get("active", True) and "U" in (dcfg.get("columns") or [])
    ]
    if _u_pool:
        rules.setdefault("U", {}).setdefault("balance", True)
        rules["U"].setdefault("balance_weight", 200)

    return cfg


def _strip_festivi_unavailability(
    unav_map: Dict[str, Dict[dt.date, Set[str]]],
    tf_fixed: List[dict],
) -> None:
    """Remove unavailability entries that would block pre-assigned festivo shifts.

    If a doctor declares unavailability on a day/shift where they've been assigned
    via sorteggio (turni_festivi.yml), the sorteggio takes precedence and the
    conflicting unavailability entry is silently ignored.
    Modifies unav_map in place.
    """
    # Shifts blocked by each column assignment
    COL_TO_SHIFTS: Dict[str, Set[str]] = {
        "D": {"Mattina", "Diurno", "Tutto il giorno", "Any"},
        "H": {"Pomeriggio", "Diurno", "Tutto il giorno", "Any"},
        "J": {"Notte", "Tutto il giorno", "Any"},
    }
    for fa in tf_fixed:
        doc = norm_name(str(fa.get("doctor", "")).strip())
        col = str(fa.get("column", "")).strip().upper()
        date_str = str(fa.get("date", "")).strip()
        try:
            date = dt.date.fromisoformat(date_str)
        except Exception:
            continue
        blocked_shifts = COL_TO_SHIFTS.get(col, set())
        if not blocked_shifts:
            continue
        doc_unav = unav_map.get(doc)
        if doc_unav and date in doc_unav:
            doc_unav[date] -= blocked_shifts
            if not doc_unav[date]:
                del doc_unav[date]


# Mapping shift name → Excel column letter for fixed-assignment injection
_FESTIVI_SHIFT_TO_COL = {
    "mattina": "D",    # festivo → slot DE (columns D+E); regular → slot D
    "pomeriggio": "H", # festivo → slot HI (columns H+I); regular → slot H
    "notte": "J",      # always → slot J
}


def load_turni_festivi(base_dir: Optional[Path] = None) -> dict:
    """Load pre-assigned holiday shifts from data/turni_festivi.yml.

    Returns a dict:
      'festivi_extra'     : list of ISO date strings to merge into cfg['festivi_extra']
      'fixed_assignments' : list of {doctor, date, column} ready for the solver
    """
    if base_dir is None:
        base_dir = Path(__file__).resolve().parent
    path = base_dir / "data" / "turni_festivi.yml"
    if not path.exists():
        return {"festivi_extra": [], "fixed_assignments": []}

    try:
        with path.open("r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except Exception:
        return {"festivi_extra": [], "fixed_assignments": []}

    if not isinstance(data, dict):
        return {"festivi_extra": [], "fixed_assignments": []}

    festivi_extra = [str(x).strip() for x in (data.get("festivi_extra") or []) if x]

    fixed: List[dict] = []
    for entry in (data.get("entries") or []):
        if not isinstance(entry, dict):
            continue
        try:
            date_str = str(entry.get("date", "")).strip()
            doctor = str(entry.get("doctor", "")).strip()
            shift_raw = str(entry.get("shift", "")).strip().lower()
            # explicit 'column' field overrides shift-to-column mapping
            col = entry.get("column") or _FESTIVI_SHIFT_TO_COL.get(shift_raw)
            if not date_str or not doctor or not col:
                continue
            dt.date.fromisoformat(date_str)  # validate format
            fixed.append({
                "doctor": doctor,
                "date": date_str,
                "column": str(col).strip().upper(),
            })
        except Exception:
            continue

    return {"festivi_extra": festivi_extra, "fixed_assignments": fixed}


def create_month_template_xlsx(
    rules_yml: "Path | str",
    year: int,
    month: int,
    out_path: "Path | str",
    sheet_name: Optional[str] = None,
) -> Path:
    """Create a minimal Excel template for the given month.

    The generated template is compatible with this generator:
    - Column A (from row 2): dates
    - Optional headers in row 1
    - Creates the worksheet specified by `sheet_name` (or a default)
    - Writes headers for columns declared in the YAML `columns:` mapping
    - Ensures `keep_empty_columns:` exist (as blank headers) for layout

    Parameters
    ----------
    rules_yml: YAML rules file path
    year, month: target year/month
    out_path: output .xlsx path
    sheet_name: worksheet name to create

    Returns
    -------
    Path to the created template.
    """
    rules_path = Path(rules_yml)
    outp = Path(out_path)
    cfg = load_rules(rules_path)

    # Determine columns from YAML
    cols_map = cfg.get("columns") or {}
    if not isinstance(cols_map, dict):
        cols_map = {}
    keep_empty = cfg.get("keep_empty_columns") or []
    if not isinstance(keep_empty, list):
        keep_empty = []

    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    if sheet_name:
        ws.title = str(sheet_name)
    else:
        ws.title = f"GUARDIE_{year}_{month:02d}"

    # Header cells (kept blank like the official model)
    _ = ws["A1"]
    _ = ws["B1"]

    # Column headers from YAML mapping (letters -> labels)
    for col_letter, label in cols_map.items():
        col_letter = str(col_letter).strip().upper()
        if not col_letter:
            continue
        ws[f"{col_letter}1"] = str(label) if label is not None else ""

    # Keep empty spacer columns
    for col_letter in keep_empty:
        col_letter = str(col_letter).strip().upper()
        if not col_letter:
            continue
        # Ensure the cell exists (leave blank)
        _ = ws[f"{col_letter}1"]

    # Fill dates
    last_day = calendar.monthrange(int(year), int(month))[1]
    r = 2
    for day in range(1, last_day + 1):
        d = dt.date(int(year), int(month), int(day))
        ws.cell(row=r, column=1).value = d
        ws.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        # Italian day labels in the Excel view (logic uses DOW_MAP internally)
        ws.cell(row=r, column=2).value = ["lunedi","martedi","mercoledi","giovedi","venerdi","sabato","domenica"][d.weekday()]
        r += 1

    # Nice-to-have formatting
    try:
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 6
    except Exception:
        pass

    # Apply model styling (if Style_Template.xlsx is present)
    _apply_model_style_to_template(ws, cfg, int(year), int(month), int(last_day))

    outp.parent.mkdir(parents=True, exist_ok=True)
    wb.save(outp)
    return outp


def create_period_template_xlsx(
    rules_yml: "Path | str",
    start_date: dt.date,
    end_date: dt.date,
    out_path: "Path | str",
    sheet_name: Optional[str] = None,
) -> Path:
    """Create an Excel template for an arbitrary inclusive date range."""
    if end_date < start_date:
        raise ValueError("end_date must be on or after start_date.")

    rules_path = Path(rules_yml)
    outp = Path(out_path)
    cfg = load_rules(rules_path)

    cols_map = cfg.get("columns") or {}
    if not isinstance(cols_map, dict):
        cols_map = {}
    keep_empty = cfg.get("keep_empty_columns") or []
    if not isinstance(keep_empty, list):
        keep_empty = []

    wb = openpyxl.Workbook()
    ws = wb.active
    if sheet_name:
        ws.title = str(sheet_name)
    else:
        ws.title = f"GUARDIE_{start_date:%Y%m%d}_{end_date:%Y%m%d}"

    _ = ws["A1"]
    _ = ws["B1"]
    for col_letter, label in cols_map.items():
        col_letter = str(col_letter).strip().upper()
        if col_letter:
            ws[f"{col_letter}1"] = str(label) if label is not None else ""
    for col_letter in keep_empty:
        col_letter = str(col_letter).strip().upper()
        if col_letter:
            _ = ws[f"{col_letter}1"]

    r = 2
    cur = start_date
    while cur <= end_date:
        ws.cell(row=r, column=1).value = cur
        ws.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=2).value = ["lunedi", "martedi", "mercoledi", "giovedi", "venerdi", "sabato", "domenica"][cur.weekday()]
        cur += dt.timedelta(days=1)
        r += 1

    try:
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 6
    except Exception:
        pass

    _apply_model_style_to_template(
        ws,
        cfg,
        int(start_date.year),
        int(start_date.month),
        (end_date - start_date).days + 1,
    )

    outp.parent.mkdir(parents=True, exist_ok=True)
    wb.save(outp)
    return outp


def load_template_days(xlsx_path: Path, sheet_name: Optional[str]=None) -> Tuple[openpyxl.Workbook, openpyxl.worksheet.worksheet.Worksheet, List[DayRow]]:
    wb = openpyxl.load_workbook(xlsx_path)
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            available = ", ".join(wb.sheetnames)
            raise KeyError(f"Worksheet '{sheet_name}' does not exist. Available: {available}")
        ws = wb[sheet_name]
    else:
        ws = wb.active
    # Find day rows: column A contains dates
    days: List[DayRow] = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, (dt.datetime, dt.date)):
            d = v.date() if isinstance(v, dt.datetime) else v
            dow = DOW_MAP[d.weekday()]
            days.append(DayRow(date=d, dow=dow, row_idx=r))
    if not days:
        raise ValueError("No date rows found in column A (starting from row 2).")
    return wb, ws, days
def load_unavailability(unav_path: Optional[Path]) -> Dict[str, Dict[dt.date, Set[str]]]:
    """
    Returns: unav[doctor][date] = {'Mattina','Pomeriggio','Notte'} (oppure 'Any' per full-day)
    """
    unav: Dict[str, Dict[dt.date, Set[str]]] = defaultdict(lambda: defaultdict(set))
    if unav_path is None:
        return unav
    if not unav_path.exists():
        raise FileNotFoundError(unav_path)
    if unav_path.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(unav_path)
    elif unav_path.suffix.lower() in [".csv", ".tsv"]:
        sep = "\t" if unav_path.suffix.lower() == ".tsv" else ","
        df = pd.read_csv(unav_path, sep=sep)
    else:
        raise ValueError("Unavailability file must be .xlsx/.xls/.csv/.tsv")
    # Flexible column names
    cols = {c.lower().strip(): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None
    c_med = pick("medico", "doctor", "name")
    c_dat = pick("data", "date", "giorno")
    c_fas = pick("fascia", "shift", "turno")
    if not (c_med and c_dat and c_fas):
        raise ValueError("Unavailability file must contain columns: Medico, Data, Fascia")
    for _, row in df.iterrows():
        med = row.get(c_med)
        dat = row.get(c_dat)
        fas = row.get(c_fas)
        if pd.isna(med) or pd.isna(dat) or pd.isna(fas):
            continue
        doctor = norm_name(med)
        date = parse_date(dat)
        for shift in shifts_from_fascia(fas):
            unav[doctor][date].add(shift)
    return unav
def collect_doctors(cfg: dict) -> List[str]:
    """
    Union of all pools/allowed in YAML, minus absolute_exclusions.
    Keeps special placeholder 'Recupero' if present.
    """
    doctors: Set[str] = set()
    # From unavailability section too
    for d in (cfg.get("unavailability") or {}).keys():
        doctors.add(norm_name(d))
    rules = cfg.get("rules", {})
    if isinstance(rules, dict):
        for _, rule in rules.items():
            if not isinstance(rule, dict):
                continue
            for k in ["allowed", "pool", "pool_other", "pool_mon_fri", "other_pool", "fallback_pool", "distribution_pool"]:
                if k in rule and isinstance(rule[k], list):
                    doctors |= {norm_name(x) for x in rule[k]}
            for k in ["fixed", "tuesday_fixed", "prefer"]:
                if k in rule and rule[k]:
                    doctors.add(norm_name(rule[k]))
    # Remove absolute exclusions
    abs_excl = {norm_name(x) for x in (cfg.get("absolute_exclusions") or [])}
    doctors = {d for d in doctors if d not in abs_excl}
    # Stable sorting: keep 'Recupero' last-ish
    doctors_list = sorted(doctors, key=lambda s: (s == "Recupero", s.lower()))
    return doctors_list
# -------------------------
# Build slots from rules
# -------------------------
def is_festivo(day: DayRow, cfg: dict) -> bool:
    extra = set()
    for x in cfg.get("festivi_extra", []) or []:
        try:
            extra.add(parse_date(x))
        except Exception:
            pass
    # Treat Sunday and Italian national holidays as "festivo".
    # Extra holidays can be provided via cfg.festivi_extra.
    hol = italy_public_holidays(int(day.date.year))
    return day.dow == "Sun" or day.date in hol or day.date in extra
def apply_unavailability(allowed: List[str], day: DayRow, shift: str, unav: Dict[str, Dict[dt.date, Set[str]]]) -> List[str]:
    out = []
    for doc in allowed:
        d_unav = unav.get(doc, {}).get(day.date, set())
        # If any shift marked 'Any' treat as full-day
        if "Any" in d_unav:
            continue
        if shift == "Any":
            if d_unav:
                continue
        elif shift in d_unav:
            continue
        out.append(doc)
    return out


def _never_in_j_set(cfg: dict) -> Set[str]:
    """Doctors that must never be assigned to night J, even via UI overrides."""
    rules = cfg.get("rules") or {}
    rJ = rules.get("J", {}) if isinstance(rules.get("J", {}), dict) else {}
    return {norm_name(d) for d in (rJ.get("never_in_J") or ["De Gregorio", "Manganaro"])}


def _fixed_assignment_allowed(cfg: dict, col: str, doctor: str) -> Tuple[bool, str]:
    """Return whether an admin fixed assignment may expand the slot domain."""
    col = str(col or "").strip().upper()
    doc = norm_name(doctor)
    if col == "J" and doc in _never_in_j_set(cfg):
        return False, f"{doc} e' in J.never_in_J"
    return True, ""


def slots_for_month(cfg: dict, days: List[DayRow], unav: Dict[str, Dict[dt.date, Set[str]]], fixed_assignments: Optional[List[dict]] = None, v_double_overrides: Optional[List[str]] = None, j_blank_week_overrides: Optional[Dict[str, List[str]]] = None, disabled_columns: Optional[List[str]] = None) -> List[Slot]:
    """
    Converts YAML column rules into per-day slots.
    Handles exception days (festivi) by merging D+E and H+I, and merging E+G always.
    """
    rules = cfg.get("rules", {})
    if not isinstance(rules, dict):
        raise ValueError("cfg.rules must be a mapping.")
    disabled_cols = {
        str(c).strip().upper()
        for c in (disabled_columns or cfg.get("disabled_columns") or [])
        if str(c).strip()
    }
    doctors_all = collect_doctors(cfg)
    doctors_set = set(doctors_all)
    # Relief valves (optional): allow specific columns to be left blank with penalties (used only if needed).
    gc = cfg.get("global_constraints") or {}
    relief = gc.get("relief_valves") or {}
    blank_penalties: Dict[str, int] = {}
    if isinstance(relief.get("allow_blank_columns"), dict):
        for _k, _v in (relief.get("allow_blank_columns") or {}).items():
            try:
                blank_penalties[str(_k).strip().upper()] = int(_v)
            except Exception:
                pass
    optional_blank_floor: Dict[str, int] = {
        # These columns may stay blank as a relief valve, but the solver should
        # first try hard to reshuffle flexible assignments (for example H -> K/L).
        "L": 20_000_000,
        "R": 10_000_000,
        "Z": 10_000_000,
    }

    def req_and_blank(col_letter: str) -> Tuple[bool, int]:
        col_letter = str(col_letter).strip().upper()
        if col_letter in blank_penalties:
            return False, max(blank_penalties[col_letter], optional_blank_floor.get(col_letter, 0))
        return True, 0
    # YAML also has inline date-only unavailability (full-day)
    for doc, dates in (cfg.get("unavailability") or {}).items():
        for ds in dates or []:
            try:
                unav[norm_name(doc)][parse_date(ds)].add("Any")
            except Exception:
                pass
    # FASE 0 — PRE-PROCESSA i fixed_assignments PRIMA della costruzione degli slot.
    # I fixed_assignment in J rendono quel medico di fatto "non disponibile" per D/F
    # lo stesso giorno (night_off same_day). Li trattiamo come indisponibilità temporanee
    # per la costruzione dell'allowed D/F.
    forced_j_by_date: Dict[dt.date, Set[str]] = {}
    forced_de_by_date: Dict[dt.date, str] = {}  # sorteggio DE (Mattina) festivi
    forced_hi_by_date: Dict[dt.date, str] = {}  # sorteggio HI (Pomeriggio) festivi
    for fa in (fixed_assignments or []):
        fa_col = str(fa.get("column","")).strip().upper()
        try:
            fa_date = dt.date.fromisoformat(str(fa.get("date","")).strip())
            fa_doc = norm_name(str(fa.get("doctor","")).strip())
            if fa_col == "J":
                ok_fixed, _reason = _fixed_assignment_allowed(cfg, fa_col, fa_doc)
                if not ok_fixed:
                    continue
                forced_j_by_date.setdefault(fa_date, set()).add(fa_doc)
            elif fa_col == "D" and fa_doc in doctors_set:
                # Il medico sorteggiato potrebbe non essere nel pool Festivi (es. Grimaldi, Calabrò)
                # → lo registriamo qui per usarlo come unico allowed nel slot DE,
                #   evitando il conflitto sum(allowed)==0 == 1 nel CP-SAT.
                forced_de_by_date[fa_date] = fa_doc
            elif fa_col == "H" and fa_doc in doctors_set:
                forced_hi_by_date[fa_date] = fa_doc
        except Exception:
            pass

    # PRE-PROCESSA j_blank_week_overrides: per settimana, quali giorni hanno J vuota.
    # Formato chiave: "YYYY-WNN" (es. "2026-W16"), valore: lista di date ISO (può essere vuota = nessun vuoto).
    # Per retrocompatibilità accetta anche una singola stringa (o None) come valore.
    _j_week_ov: Dict[tuple, Set[dt.date]] = {}
    for _wk_str, _bd_val in (j_blank_week_overrides or {}).items():
        try:
            _parts = str(_wk_str).split("-W")
            _iso_key = (int(_parts[0]), int(_parts[1]))
            if isinstance(_bd_val, (list, tuple, set)):
                _blank_dates = {dt.date.fromisoformat(str(x).strip()) for x in _bd_val if str(x).strip()}
            elif _bd_val:
                _blank_dates = {dt.date.fromisoformat(str(_bd_val).strip())}
            else:
                _blank_dates = set()
            _j_week_ov[_iso_key] = _blank_dates
        except Exception:
            pass

    # PRE-PROCESSA v_double_overrides: date esatte in cui V è in doppio (al posto del venerdì)
    # Se quella settimana ha un override, il venerdì di quella settimana diventa turno singolo.
    # Sentinel "NODOUBLE:{year}:{week}" → quella settimana non ha nessun turno doppio.
    _v_override_dates: Set[dt.date] = set()
    _v_no_double_weeks: Set[tuple] = set()
    if v_double_overrides:
        for _ds in v_double_overrides:
            _ds = str(_ds).strip()
            if _ds.startswith("NODOUBLE:"):
                try:
                    _, _yr_wk = _ds.split(":", 1)
                    _yr_s, _wk_s = _yr_wk.split(":")
                    _v_no_double_weeks.add((int(_yr_s), int(_wk_s)))
                except Exception:
                    pass
            else:
                try:
                    _v_override_dates.add(dt.date.fromisoformat(_ds))
                except Exception:
                    pass
    # ISO week keys delle settimane che hanno un override (doppio spostato O nessun doppio)
    _v_override_weeks: Set[tuple] = {d.isocalendar()[:2] for d in _v_override_dates} | _v_no_double_weeks

    slots: List[Slot] = []
    def mk_allowed(pool: List[str]) -> List[str]:
        # keep only known doctors + keep 'Recupero' if used
        out = [norm_name(x) for x in pool if norm_name(x) in doctors_set]
        return out
    for day in days:
        festivo = is_festivo(day, cfg)
        # Optional: on Saturdays, assign the SAME doctor to K and T (single combined slot K+T)
        gc = cfg.get("global_constraints", {}) or {}
        merge_KT_sat = bool(gc.get("saturday_K_equals_T", False)) and (day.dow == "Sat") and (not festivo)             and ("K" in rules) and ("T" in rules) and dayspec_contains(day.dow, (rules.get("T") or {}).get("days"))
        # ---- C: Reperibilità (daily)
        if "C_reperibilita" in rules:
            r = rules["C_reperibilita"]
            excluded = {norm_name(x) for x in (r.get("excluded") or [])}
            pool = [d for d in doctors_all if d not in excluded and d != "Recupero"]  # usually a real doctor
            pool = apply_unavailability(pool, day, "Any", unav)
            slots.append(Slot(day, f"{day.date}-C", ["C"], pool, required=True, shift="Any", rule_tag="C_reperibilita"))
        # ---- Morning D/F and E/G and Afternoon H/I depend on festivo
        if festivo:
            rFest = rules.get("Festivi", {}) if isinstance(rules.get("Festivi", {}), dict) else {}
            fest_excl = {norm_name(x) for x in (rFest.get("excluded") or [])}

            # DE unified (D+E) – required
            # Se c'è un medico sorteggiato (anche escluso dal pool Festivi, es. Grimaldi/Calabrò),
            # lo mettiamo come UNICO allowed per evitare sum(allowed)==0==1 nel CP-SAT.
            _forced_de = forced_de_by_date.get(day.date)
            if _forced_de:
                allowed_de = [_forced_de]
            else:
                fest_pool_m = rFest.get("pool_mattina") or rFest.get("pool") or []
                allowed_de = mk_allowed(fest_pool_m)
                if not allowed_de:
                    allowed_de = [d for d in doctors_all if d != "Recupero" and d not in fest_excl]
                else:
                    allowed_de = [d for d in allowed_de if d not in fest_excl and d != "Recupero"]
                allowed_de = apply_unavailability(allowed_de, day, "Mattina", unav)
            slots.append(Slot(day, f"{day.date}-DE", ["D","E"], allowed_de, required=True, shift="Mattina", rule_tag="Festivo_DE"))

            # HI unified (H+I) – required
            _forced_hi = forced_hi_by_date.get(day.date)
            if _forced_hi:
                allowed_hi = [_forced_hi]
            else:
                fest_pool_p = rFest.get("pool_pomeriggio") or rFest.get("pool") or []
                allowed_hi = mk_allowed(fest_pool_p)
                if not allowed_hi:
                    allowed_hi = [d for d in doctors_all if d != "Recupero" and d not in fest_excl]
                else:
                    allowed_hi = [d for d in allowed_hi if d not in fest_excl and d != "Recupero"]
                allowed_hi = apply_unavailability(allowed_hi, day, "Pomeriggio", unav)
            slots.append(Slot(day, f"{day.date}-HI", ["H","I"], allowed_hi, required=True, shift="Pomeriggio", rule_tag="Festivo_HI"))
        else:
            # D / F (Mon-Sat)
            if "D_F" in rules and dayspec_contains(day.dow, rules["D_F"].get("days")):
                r = rules["D_F"]
                pair_docs = mk_allowed(r.get("allowed") or [])
                pair_avail = apply_unavailability(pair_docs, day, "Mattina", unav)
                # Rimuovi i medici forzati in J quel giorno (night_off same_day li esclude da D/F)
                forced_j_today = forced_j_by_date.get(day.date, set())
                pair_avail = [d for d in pair_avail if d not in forced_j_today]

                # Fallback source = H.pool_mon_fri (as requested)
                h_rule = rules.get("H", {}) if isinstance(rules.get("H", {}), dict) else {}
                h_pool = mk_allowed(h_rule.get("pool_mon_fri") or [])
                h_avail = apply_unavailability(h_pool, day, "Mattina", unav)

                # Ultimate fallback: any doctor available that morning (MAI Recupero)
                _recupero_n = norm_name("Recupero")
                any_pool = apply_unavailability(
                    [d for d in sorted(doctors_set) if d != _recupero_n], day, "Mattina", unav
                )

                # Se solo uno del pair disponibile → share obbligatorio (lui fa D e F)
                # Se nessuno del pair → fallback ordinato:
                #   1. H pool disponibile, bilanciato/stabilizzato dal solver
                #   2. qualsiasi medico disponibile solo se il pool H non può coprire
                # Se entrambi → solo il pair, no share
                if len(pair_avail) == 1:
                    allowed_df = pair_avail
                    prefer_share = True
                elif len(pair_avail) == 0:
                    allowed_df = h_avail or any_pool or sorted(doctors_set)
                    prefer_share = True
                else:
                    allowed_df = pair_avail
                    prefer_share = False

                slots.append(Slot(day, f"{day.date}-D", ["D"], allowed_df, required=True, shift="Mattina", rule_tag="D_F.D", force_same_doctor=prefer_share))
                slots.append(Slot(day, f"{day.date}-F", ["F"], allowed_df, required=True, shift="Mattina", rule_tag="D_F.F", force_same_doctor=prefer_share))
# EG paired (Mon-Sat)
            if "E_G" in rules and dayspec_contains(day.dow, rules["E_G"].get("days")):
                r = rules["E_G"]
                allowed = mk_allowed(r.get("allowed") or [])
                allowed = apply_unavailability(allowed, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-EG", ["E","G"], allowed, required=True, shift="Mattina", rule_tag="E_G"))
            # H (Mon-Sat) + I (Mon-Sat)
            # MODIFICA 1: Grimaldi e Calabrò NON devono MAI comparire in H
            if "H" in rules:
                rH = rules["H"]
                # Saturdays must be covered; for Mon-Fri, use pool_mon_fri only
                if day.dow == "Sat" or dayspec_contains(day.dow, "Mon-Fri"):
                    allowed = mk_allowed(rH.get("pool_mon_fri") or [])
                    # Grimaldi e Calabrò esclusi esplicitamente da H
                    _h_excl = {norm_name("Grimaldi"), norm_name("Calabrò")}
                    allowed = [d for d in allowed if norm_name(d) not in _h_excl]
                    allowed = apply_unavailability(allowed, day, "Pomeriggio", unav)
                    slots.append(Slot(day, f"{day.date}-H", ["H"], allowed, required=True, shift="Pomeriggio", rule_tag="H"))
            if "I" in rules and dayspec_contains(day.dow, "Mon-Sat"):
                rI = rules["I"]
                allowed = mk_allowed(rI.get("distribution_pool") or [])
                allowed = apply_unavailability(allowed, day, "Pomeriggio", unav)
                # In practice I is an afternoon activity; required Mon-Sat
                slots.append(Slot(day, f"{day.date}-I", ["I"], allowed, required=True, shift="Pomeriggio", rule_tag="I"))
        # ---- Night J (daily except Thu if configured; MODIFICA 7: override data specifica)
        if "J" in rules:
            rJ = rules["J"]
            # Data di override (es. "2026-03-04" per mercoledì 4 marzo al posto di giovedì 5)
            _week_key_j = day.date.isocalendar()[:2]
            if _week_key_j in _j_week_ov:
                # Override UI: questa settimana ha un'impostazione specifica (uno o più giorni vuoti)
                _skip_j = day.date in _j_week_ov[_week_key_j]
            else:
                # Comportamento default: giovedì vuoto + eventuale override YAML
                _j_override_raw = (cfg.get("global_constraints") or {}).get("j_blank_override_date")
                _j_override_date = None
                if _j_override_raw:
                    try:
                        _j_override_date = dt.date.fromisoformat(str(_j_override_raw))
                    except Exception:
                        pass
                _is_normal_thu_blank = rJ.get("thursday_blank", False) and day.dow == "Thu"
                _is_override_blank = (_j_override_date is not None and day.date == _j_override_date)
                _thu_suppressed_by_override = False
                if _j_override_date is not None and day.dow == "Thu":
                    import datetime as _dt3
                    _thu_week_mon = day.date - _dt3.timedelta(days=day.date.weekday())
                    _ov_week_mon = _j_override_date - _dt3.timedelta(days=_j_override_date.weekday())
                    if _thu_week_mon == _ov_week_mon:
                        _thu_suppressed_by_override = True
                _skip_j = (_is_normal_thu_blank and not _thu_suppressed_by_override) or _is_override_blank
            if not _skip_j:
                # Se esiste un fixed_assignment in J per questo giorno,
                # lo slot usa SOLO quel medico — l'indisponibilità viene ignorata
                # perché l'admin ha deciso esplicitamente.
                forced_j_today = forced_j_by_date.get(day.date, set())
                if forced_j_today:
                    allowed = [d for d in forced_j_today if d in doctors_set]
                else:
                    allowed = mk_allowed(rJ.get("pool_other") or [])
                    # add quota doctors even if not in pool_other
                    for special in (rJ.get("monthly_quotas") or {}).keys():
                        if special in doctors_set and special not in allowed:
                            allowed.append(special)
                    allowed = [d for d in allowed if d != "Recupero"]
                    # Esclusioni permanenti da J (mai in notte, qualunque giorno)
                    j_never = {norm_name(d) for d in (rJ.get("never_in_J") or ["De Gregorio", "Manganaro"])}
                    allowed = [d for d in allowed if norm_name(d) not in j_never]
                    # Weekend exclusions (e.g., Calabrò not allowed on Sat/Sun nights)
                    wex = [norm_name(x) for x in (rJ.get('weekend_excluded_doctors') or [])]
                    if day.dow in ['Sat','Sun'] and wex:
                        allowed = [d for d in allowed if norm_name(d) not in set(wex)]
                    # festivi_notti filter: rimuove medici esclusi da J nei giorni festivi
                    if is_festivo(day, cfg):
                        _fest_notti_excl = cfg.get("pool_festivi_notti_excluded") or set()
                        if _fest_notti_excl:
                            allowed = [d for d in allowed if norm_name(d) not in _fest_notti_excl]
                    allowed = apply_unavailability(allowed, day, "Notte", unav)
                slots.append(Slot(day, f"{day.date}-J", ["J"], allowed, required=True, shift="Notte", rule_tag="J"))
        # ---- K Letto (daily, but blank on Sundays/festivi)
        # If merge_KT_sat is enabled, we create a single combined slot K+T on Saturday.
        if merge_KT_sat:
            rK = rules.get("K", {}) if isinstance(rules.get("K", {}), dict) else {}
            rT = rules.get("T", {}) if isinstance(rules.get("T", {}), dict) else {}
            allowed_k = mk_allowed(rK.get("pool") or [])
            allowed_t = mk_allowed(rT.get("pool") or [])
            allowed_k = apply_unavailability(allowed_k, day, "Mattina", unav)
            allowed_t = apply_unavailability(allowed_t, day, "Mattina", unav)
            inter = [d for d in allowed_k if d in set(allowed_t)]
            allowed = inter if inter else sorted({*allowed_k, *allowed_t})
            slots.append(Slot(day, f"{day.date}-KT", ["K","T"], allowed, required=True, shift="Mattina", rule_tag="K_T_SAT"))
        elif "K" in rules and not festivo:
            rK = rules["K"]
            allowed = mk_allowed(rK.get("pool") or [])
            allowed = apply_unavailability(allowed, day, "Mattina", unav)
            slots.append(Slot(day, f"{day.date}-K", ["K"], allowed, required=True, shift="Mattina", rule_tag="K"))
        # ---- L Padiglioni (Mon-Wed)
        if "L" in rules and not festivo:
            rL = rules["L"]
            if dayspec_contains(day.dow, rL.get("days")):
                pool = mk_allowed(rL.get("pool_other") or [])
                # allow Recupero as placeholder
                if "Recupero" in doctors_set and "Recupero" not in pool:
                    pool.append("Recupero")
                pool = apply_unavailability(pool, day, "Mattina", unav)
                # L usa sempre il relief valve (20K) — priorità inferiore a H (5M obbligatorio)
                req_relief, bp = req_and_blank("L")
                slots.append(Slot(day, f"{day.date}-L", ["L"], pool, required=req_relief, blank_penalty=bp, shift="Mattina", rule_tag="L"))
        # ---- Q Eco base (Mon-Sat)
        if "Q" in rules and not festivo:
            r = rules["Q"]
            if dayspec_contains(day.dow, r.get("days")):
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-Q", ["Q"], pool, required=True, shift="Mattina", rule_tag="Q"))
        # ---- R (Mon-Fri)
        if "R" in rules and not festivo:
            r = rules["R"]
            if dayspec_contains(day.dow, r.get("days")):
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                # R usa sempre il relief valve (40K) — priorità inferiore agli slot obbligatori
                req_relief, bp = req_and_blank("R")
                slots.append(Slot(day, f"{day.date}-R", ["R"], pool, required=req_relief, blank_penalty=bp, shift="Mattina", rule_tag="R"))
        # ---- S (Wed, optional if can be absorbed in R)
        if "S" in rules and not festivo:
            r = rules["S"]
            if dayspec_contains(day.dow, r.get("days") or r.get("day")):
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                required = not bool(r.get("if_not_dedicated_put_in_R", False))
                slots.append(Slot(day, f"{day.date}-S", ["S"], pool, required=required, shift="Mattina", rule_tag="S"))
        # ---- T Interni (Mon-Sat)
        if "T" in rules and not merge_KT_sat and not festivo:
            r = rules["T"]
            if dayspec_contains(day.dow, r.get("days")):
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-T", ["T"], pool, required=True, shift="Mattina", rule_tag="T"))
        # ---- U Contr.PM (Mon-Tue)
        if "U" in rules and not festivo:
            r = rules["U"]
            if dayspec_contains(day.dow, r.get("days")):
                # Lunedì Contr.PM è pomeridiano; martedì è mattutino
                u_shift = "Pomeriggio" if day.dow == "Mon" else "Mattina"
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, u_shift, unav)
                # MODIFICA 3: se lunedì e V ha solo Allegra disponibile (o è assegnato ad Allegra),
                # U deve essere SOLO Crea o Dattilo.
                if day.dow == "Mon" and r.get("v_allegra_monday_constraint", False):
                    rV = rules.get("V", {})
                    v_pool_avail = apply_unavailability(mk_allowed(rV.get("pool") or []), day, "Pomeriggio", unav)
                    _allegra = norm_name("Allegra")
                    _crea = norm_name("Crea")
                    _dattilo = norm_name("Dattilo")
                    # Se Allegra è l'unico disponibile in V (o i soli disponibili sono Allegra),
                    # oppure Crea e Dattilo non sono nel pool V → forza U = {Crea, Dattilo}
                    only_allegra_in_v = all(norm_name(d) == _allegra for d in v_pool_avail) if v_pool_avail else False
                    if only_allegra_in_v:
                        restricted = [d for d in pool if norm_name(d) in {_crea, _dattilo}]
                        if restricted:  # applica solo se almeno uno è disponibile
                            pool = restricted
                slots.append(Slot(day, f"{day.date}-U", ["U"], pool, required=True, shift=u_shift, rule_tag="U"))
        # ---- V Sala PM (Mon,Wed,Fri)
        # Default: venerdì = turno doppio (Crea + Dattilo|Allegra), lun/mer = singolo.
        # Override admin: se un lunedì o mercoledì è in _v_override_dates, quella settimana
        # il doppio è su quel giorno e il venerdì diventa singolo.
        if "V" in rules and not festivo:
            r = rules["V"]
            if dayspec_contains(day.dow, r.get("days")):
                # Lunedì la Sala PM è di pomeriggio; mercoledì e venerdì è di mattina
                v_shift = "Pomeriggio" if day.dow == "Mon" else "Mattina"
                pool_base = mk_allowed(r.get("pool") or [])
                pool_base = apply_unavailability(pool_base, day, v_shift, unav)
                # Determina se questo giorno è il "turno doppio" della settimana
                _week_key = day.date.isocalendar()[:2]
                _is_override_double = day.date in _v_override_dates
                _is_default_double = (day.dow == "Fri") and (_week_key not in _v_override_weeks)
                _is_double_day = _is_override_double or _is_default_double
                if _is_double_day:
                    crea = norm_name(r.get("friday_required_doctor") or "Crea")
                    pool_crea = [crea] if crea in pool_base else []
                    other_allowed = {norm_name("Dattilo"), norm_name("Allegra")}
                    pool_other = [d for d in pool_base if norm_name(d) in other_allowed and norm_name(d) != crea]
                    # Turno doppio: solo se entrambi i pool sono non vuoti
                    if (not pool_crea) or (not pool_other):
                        pool_crea = []
                        pool_other = []
                    slots.append(Slot(day, f"{day.date}-V1", ["V"], pool_crea, required=True, shift=v_shift, rule_tag="V"))
                    slots.append(Slot(day, f"{day.date}-V2", ["V"], pool_other, required=True, shift=v_shift, rule_tag="V"))
                else:
                    slots.append(Slot(day, f"{day.date}-V", ["V"], pool_base, required=True, shift=v_shift, rule_tag="V"))
        # ---- Z Vascolare (Wed,Fri)
        if "Z" in rules and not festivo:
            r = rules["Z"]
            if dayspec_contains(day.dow, r.get("days")):
                pool = mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                # Z usa il relief valve (30K) — si sacrifica prima di R (40K)
                req_relief_z, bp_z = req_and_blank("Z")
                slots.append(Slot(day, f"{day.date}-Z", ["Z"], pool, required=req_relief_z, blank_penalty=bp_z, shift="Mattina", rule_tag="Z"))
        # ---- W Ergometria/CPET (Mon-Fri; Tue fixed) — escluso nei festivi
        if "W" in rules and not festivo:
            r = rules["W"]
            if day.dow in ["Mon","Tue","Wed","Thu","Fri"]:
                if day.dow == "Tue" and r.get("tuesday_fixed"):
                    fixed = norm_name(r["tuesday_fixed"])
                    pool = [fixed] if fixed in doctors_set else []
                else:
                    pool = mk_allowed(r.get("other_days_pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                # W è sempre richiesto (lun-ven); se pool vuoto per indisponibilità
                # lo slot diventa opzionale con blank_penalty alta → cella gialla
                if pool:
                    required = True
                    blank_pen = 0
                else:
                    required = False
                    blank_pen = 5000
                slots.append(Slot(day, f"{day.date}-W", ["W"], pool, required=required,
                                  blank_penalty=blank_pen, shift="Mattina", rule_tag="W"))
        # ---- Y Amb specialistici (Mon only)
        # Requirement:
        #  - Every Monday: 1 doctor among other_pool
        #  - PLUS: on exactly 2 Mondays: also 'Recupero' (appended in the same cell)
        if "Y" in rules and not festivo:
            r = rules["Y"]
            if dayspec_contains(day.dow, r.get("days") or r.get("day")):
                # Main doctor (always required)
                pool_main = [d for d in mk_allowed(r.get("other_pool") or []) if d != "Recupero"]
                pool_main = apply_unavailability(pool_main, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-Y", ["Y"], pool_main, required=True, shift="Mattina", rule_tag="Y_MAIN"))
        # ---- AB Holter/Brugada/FA (Thu)
        if "AB" in rules and not festivo:
            r = rules["AB"]
            if r.get("weekly", False) and dayspec_contains(day.dow, r.get("fixed_day")):
                # MODIFICA 5: nessuna preferenza per Crea il giovedì — pool bilanciato
                prefer = norm_name(r.get("prefer") or "")
                pool = []
                if prefer and prefer in doctors_set:
                    pool.append(prefer)
                pool += mk_allowed(r.get("fallback_pool") or [])
                # unique list preserving order
                seen=set(); pool=[x for x in pool if not (x in seen or seen.add(x))]
                pool = apply_unavailability(pool, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-AB", ["AB"], pool, required=True, shift="Mattina", rule_tag="AB"))
            # Slot aggiuntivo al Sabato (2 al mese) SOLO con CREA (quota gestita nel solver)
            sat_n = int(r.get("saturday_per_month", 0) or 0)
            if sat_n > 0 and day.dow == "Sat":
                doc_sat = norm_name(r.get("saturday_only_doctor") or "Crea")
                pool_sat = [doc_sat] if doc_sat in doctors_set else []
                pool_sat = apply_unavailability(pool_sat, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-AB_SAT", ["AB"], pool_sat, required=False, shift="Mattina", rule_tag="AB_SAT"))
        # ---- AC Scintigrafia (Tue/Wed fixed)
        if "AC" in rules and not festivo:
            r = rules["AC"]
            if dayspec_contains(day.dow, r.get("days")):
                fixed = norm_name(r.get("fixed") or "")
                pool = [fixed] if fixed in doctors_set else mk_allowed(r.get("pool") or [])
                pool = apply_unavailability(pool, day, "Mattina", unav)
                slots.append(Slot(day, f"{day.date}-AC", ["AC"], pool, required=True, shift="Mattina", rule_tag="AC"))
    # ── Disattivazione esplicita colonne ────────────────────────────────────────
    # Se una colonna è disattivata dall'admin per questa generazione, non deve
    # produrre slot né fallback: la cella resta intenzionalmente vuota.
    if disabled_cols:
        _filtered_slots: List[Slot] = []
        for s in slots:
            _enabled_columns = [c for c in (s.columns or []) if str(c).strip().upper() not in disabled_cols]
            if not _enabled_columns:
                continue
            if _enabled_columns != s.columns:
                s.columns = _enabled_columns
            _filtered_slots.append(s)
        slots = _filtered_slots

    # ── Espansione pool per colonne indispensabili (critical_services) ──────────
    # Per ogni colonna marcata come "indispensabile" in pool_critical_services:
    # se il pool primario è vuoto o molto ridotto, si espande a qualsiasi medico
    # disponibile per quel turno (con penalità nel solver per scoraggiarne l'uso).
    _abs_excl_norm = {norm_name(x) for x in (cfg.get("absolute_exclusions") or [])}
    # Grimaldi e Calabrò vanno SOLO in D/F (mai in pool di emergenza di altre colonne)
    _df_only_norm = {norm_name("Grimaldi"), norm_name("Calabrò")}
    _emerg_any_pool = [d for d in doctors_all
                       if norm_name(d) not in _abs_excl_norm
                       and norm_name(d) != "Recupero"
                       and norm_name(d) not in _df_only_norm]
    _critical_svc = cfg.get("pool_critical_services") or {}

    for s in slots:
        if not s.required:
            continue
        _match_col = next((c for c in (s.columns or []) if c in _critical_svc), None)
        if not _match_col:
            continue
        _spec = _critical_svc.get(_match_col, {})
        _fb = _spec.get("fallback", "")
        if _fb == "any":
            _fb_base = _emerg_any_pool
        elif isinstance(_fb, list) and _fb:
            _fb_base = [norm_name(d) for d in _fb
                        if norm_name(d) in doctors_set and norm_name(d) != "Recupero"]
        else:
            continue
        _primary_set = set(s.allowed)
        if _primary_set:
            continue
        _fb_avail = apply_unavailability(_fb_base, s.day, s.shift, unav)
        _new_emerg = [d for d in _fb_avail if d not in _primary_set]
        if _new_emerg:
            s.allowed = s.allowed + _new_emerg
            s.emergency_doctors = _new_emerg

    # ── Emergency expansion per H e E/G ────────────────────────────────────
    # Se il pool primario non copre: qualsiasi medico disponibile eccetto
    # Recupero, Grimaldi e Calabrò (che vanno solo in D/F).
    _h_eg_excl_norm = _df_only_norm | {"Recupero"} | _abs_excl_norm
    _h_eg_emerg_base = [d for d in doctors_all if norm_name(d) not in _h_eg_excl_norm]
    for s in slots:
        if not s.required or s.rule_tag not in {"H", "E_G"}:
            continue
        _primary_set = set(s.allowed)
        if _primary_set:
            continue
        _emerg_avail = apply_unavailability(_h_eg_emerg_base, s.day, s.shift, unav)
        _new_emerg = [d for d in _emerg_avail if d not in _primary_set]
        if _new_emerg:
            s.allowed = s.allowed + _new_emerg
            existing = s.emergency_doctors or []
            s.emergency_doctors = existing + _new_emerg

    # ── Safety: rimuove Recupero da qualsiasi slot festivo (DE, HI) ─────────────
    # Recupero è un placeholder — non lavora mai nei festivi.
    _recupero_str = "Recupero"
    for s in slots:
        if s.rule_tag in {"Festivo_DE", "Festivo_HI"}:
            if _recupero_str in s.allowed:
                s.allowed = [d for d in s.allowed if d != _recupero_str]
                if s.emergency_doctors:
                    s.emergency_doctors = [d for d in s.emergency_doctors if d != _recupero_str]

    # ── Esclusione sabato diurno da GUI ────────────────────────────────────────
    # Si applica dopo i fallback, così il medico escluso non rientra come
    # emergenza. Non riguarda J notte, C reperibilita' o slot "Any".
    _sat_day_excl = {norm_name(d) for d in (cfg.get("pool_saturday_day_excluded") or set())}
    if _sat_day_excl:
        for s in slots:
            if s.day.dow != "Sat":
                continue
            if (s.shift or "") not in {"Mattina", "Pomeriggio"}:
                continue
            s.allowed = [d for d in s.allowed if norm_name(d) not in _sat_day_excl]
            if s.emergency_doctors:
                s.emergency_doctors = [
                    d for d in s.emergency_doctors
                    if norm_name(d) not in _sat_day_excl
                ]

    # ── Validate domains: slot required con pool ancora vuoto ────────────────
    # Per colonne INDISPENSABILI: già espanso sopra con emergency pool.
    # Per colonne NON indispensabili: cella gialla (slot opzionale/blank).
    # Il solver non cerca medici fuori dal pool — la cella rimane vuota
    # ed evidenziata di giallo nell'Excel di output.
    for s in slots:
        if not s.allowed:
            s.empty_domain = True
            if s.required:
                s.required = False
                s.blank_penalty = 5000  # → cella gialla nell'output Excel
    return slots
# -------------------------
# Solver (OR-Tools CP-SAT)
# -------------------------
def _max_bipartite_matching(slots_day: List[Slot]) -> Tuple[int, Dict[str, str]]:
    """
    Simple DFS-based bipartite matching: Slot -> Doctor.
    Returns:
      matched_count, slot_to_doc (by slot_id)
    """
    # Ensure deterministic iteration
    slots_day = list(slots_day)
    doc_to_slot: Dict[str, Slot] = {}
    def try_assign(slot: Slot, seen: Set[str]) -> bool:
        for d in dict.fromkeys(slot.allowed):
            if d in seen:
                continue
            seen.add(d)
            if d not in doc_to_slot or try_assign(doc_to_slot[d], seen):
                doc_to_slot[d] = slot
                return True
        return False
    for slot in sorted(slots_day, key=lambda s: (len(s.allowed), s.slot_id)):
        try_assign(slot, set())
    matched_slot_ids = set(s.slot_id for s in doc_to_slot.values())
    slot_to_doc: Dict[str, str] = {}
    for d, s in doc_to_slot.items():
        slot_to_doc[s.slot_id] = d
    return len(matched_slot_ids), slot_to_doc
def diagnose_day_level(days: List[DayRow], slots: List[Slot]) -> List[Dict]:
    """
    Day-level feasibility diagnostics (ignores cross-day constraints such as night spacing).
    Useful to pinpoint single-day bottlenecks created by unavailability.
    """
    slots_by_day: Dict[dt.date, List[Slot]] = defaultdict(list)
    for s in slots:
        slots_by_day[s.day.date].append(s)
    report: List[Dict] = []
    def _diag_slot_is_exempt_daily(s: Slot) -> bool:
        # The real model exempts C from daily uniqueness. This diagnostic is a
        # quick bipartite check, so exempt slots must not consume a unique doctor.
        return any(str(c).strip().upper() == "C" for c in (s.columns or []))

    for day in days:
        day_slots = slots_by_day.get(day.date, [])
        # consider only required slots (and also "penalized optional" slots) as "should be filled"
        must_fill = [
            s for s in day_slots
            if not _diag_slot_is_exempt_daily(s)
            and (s.required or (getattr(s, "blank_penalty", 0) and int(getattr(s, "blank_penalty", 0)) > 0))
        ]
        if not must_fill:
            continue
        matched, slot_to_doc = _max_bipartite_matching(must_fill)
        ok = (matched == len(must_fill))
        union_docs = sorted({d for s in must_fill for d in s.allowed})
        tight = sorted([(s.slot_id, s.columns, len(s.allowed)) for s in must_fill], key=lambda x: x[2])[:6]
        if not ok:
            # identify some unmatched slots
            matched_ids = set(slot_to_doc.keys())
            unmatched = []
            for s in sorted(must_fill, key=lambda s: (len(s.allowed), s.slot_id)):
                if s.slot_id not in matched_ids:
                    unmatched.append({"slot_id": s.slot_id, "columns": s.columns, "allowed_n": len(s.allowed), "allowed": s.allowed[:15]})
            report.append({
                "date": day.date.isoformat(),
                "dow": day.dow,
                "required_slots": len(must_fill),
                "union_doctors": len(union_docs),
                "tightest_slots": tight,
                "unmatched_slots": unmatched[:6],
            })
    return report
def build_relief_log(days: List[DayRow], slots: List[Slot], assignment: Dict[str, Optional[str]]) -> Dict:
    """
    Summarize where relief valves were used (K=T same doctor, blanks on penalized optional columns).
    """
    slots_by_day: Dict[dt.date, List[Slot]] = defaultdict(list)
    for s in slots:
        slots_by_day[s.day.date].append(s)
    kt_share_days: List[str] = []
    df_share_days: List[str] = []
    df_forced_same_days: List[str] = []
    blank_cols: Dict[str, List[str]] = defaultdict(list)
    for day in days:
        day_slots = slots_by_day.get(day.date, [])
        # blanks
        for s in day_slots:
            if (not s.required) and int(getattr(s, "blank_penalty", 0)) > 0:
                if assignment.get(s.slot_id) is None:
                    for c in s.columns:
                        blank_cols[str(c)].append(day.date.isoformat())
        # K/T share (only when K and T are separate slots)
        sk = next((s for s in day_slots if s.columns == ["K"]), None)
        st = next((s for s in day_slots if s.columns == ["T"]), None)
        if sk and st:
            dk = assignment.get(sk.slot_id)
            dt_ = assignment.get(st.slot_id)
            if dk is not None and dk == dt_:
                kt_share_days.append(day.date.isoformat())
        # D/F share (when fallback forces/needs D=F)
        sD = next((s for s in day_slots if s.columns == ["D"]), None)
        sF = next((s for s in day_slots if s.columns == ["F"]), None)
        if sD and sF:
            dD = assignment.get(sD.slot_id)
            dF = assignment.get(sF.slot_id)
            if dD is not None and dD == dF:
                df_share_days.append(day.date.isoformat())
                if bool(getattr(sD, "force_same_doctor", False)) and bool(getattr(sF, "force_same_doctor", False)):
                    df_forced_same_days.append(day.date.isoformat())
    return {
        "kt_share_days": kt_share_days,
        "df_share_days": df_share_days,
        "df_forced_same_days": df_forced_same_days,
        "blank_columns": dict(blank_cols),
    }

def build_daily_diagnostic(
    days: List[DayRow],
    slots: List[Slot],
    assignment: Dict[str, Optional[str]],
    cfg: dict,
) -> List[Dict]:
    """Build a per-day diagnostic: blanks, relief valves, forced blanks."""
    slots_by_day: Dict[dt.date, List[Slot]] = defaultdict(list)
    for s in slots:
        slots_by_day[s.day.date].append(s)

    DOW_ITA = {"Mon": "Lun", "Tue": "Mar", "Wed": "Mer", "Thu": "Gio",
               "Fri": "Ven", "Sat": "Sab", "Sun": "Dom"}
    diagnostics: List[Dict] = []

    for day in days:
        day_slots = slots_by_day.get(day.date, [])
        issues: List[str] = []

        for s in day_slots:
            doc = assignment.get(s.slot_id)
            cols_str = "+".join(s.columns)

            if doc is None:
                if getattr(s, "empty_domain", False):
                    issues.append(f"col {cols_str} vuota (nessun candidato eleggibile)")
                elif s.required and int(getattr(s, "blank_penalty", 0)) >= 5_000_000:
                    issues.append(f"col {cols_str} NON COPERTA (obbligatoria)")
                elif int(getattr(s, "blank_penalty", 0)) > 0:
                    bp = int(getattr(s, "blank_penalty", 0))
                    issues.append(f"col {cols_str} blank (relief valve, penalty={bp})")
                elif not s.required:
                    pass  # optional slot left blank is normal
                else:
                    issues.append(f"col {cols_str} blank (non assegnata)")

        # K/T same doctor
        sk = next((s for s in day_slots if s.columns == ["K"]), None)
        st_ = next((s for s in day_slots if s.columns == ["T"]), None)
        if sk and st_:
            dk = assignment.get(sk.slot_id)
            dt_ = assignment.get(st_.slot_id)
            if dk is not None and dk == dt_:
                issues.append(f"K=T stesso medico ({dk})")

        # D/F same doctor
        sD = next((s for s in day_slots if s.columns == ["D"]), None)
        sF = next((s for s in day_slots if s.columns == ["F"]), None)
        if sD and sF:
            dD = assignment.get(sD.slot_id)
            dF = assignment.get(sF.slot_id)
            if dD is not None and dD == dF:
                forced = getattr(sD, "force_same_doctor", False) and getattr(sF, "force_same_doctor", False)
                tag = " (forzato)" if forced else ""
                issues.append(f"D=F stesso medico ({dD}){tag}")

        if issues:
            festivo = is_festivo(day, cfg)
            diagnostics.append({
                "date": day.date.isoformat(),
                "dow": DOW_ITA.get(day.dow, day.dow),
                "festivo": festivo,
                "issues": issues,
            })

    return diagnostics

def _is_heavy_festive_slot(cfg: dict, slot: Slot) -> bool:
    """Return True for the burdens the user expects to be pooled together."""
    if slot.rule_tag in {"Festivo_DE", "Festivo_HI"}:
        return True
    if slot.columns == ["J"]:
        return slot.day.dow in {"Sat", "Sun"} or is_festivo(slot.day, cfg)
    return False


def _fixed_heavy_slot_ids(slots: List[Slot], fixed_assignments: Optional[List[dict]]) -> Set[str]:
    fixed_ids: Set[str] = set()
    if not fixed_assignments:
        return fixed_ids
    for raw in fixed_assignments:
        try:
            fdate = dt.date.fromisoformat(str(raw.get("date", "")).strip())
            fcol = str(raw.get("column", "")).strip().upper()
        except Exception:
            continue
        if not fcol:
            continue
        for s in slots:
            if s.day.date == fdate and fcol in {str(c).upper() for c in (s.columns or [])}:
                fixed_ids.add(s.slot_id)
    return fixed_ids


def _repair_heavy_festive_assignments(
    cfg: dict,
    slots: List[Slot],
    assignment: Dict[str, Optional[str]],
    fixed_assignments: Optional[List[dict]] = None,
    max_passes: int = 80,
) -> Tuple[Dict[str, Optional[str]], List[Dict]]:
    """Second-look local repair for J weekend/festive + festive D/E/H/I.

    The CP-SAT objective already carries these penalties, but large real-world
    models can still settle on a locally unpleasant distribution. This pass only
    accepts moves that improve the unified heavy-duty distribution and only
    moves a slot to a doctor already present in that slot's solved domain.
    """
    if not slots or not assignment:
        return assignment, []

    fixed_ids = _fixed_heavy_slot_ids(slots, fixed_assignments)
    heavy_slots = [
        s for s in slots
        if _is_heavy_festive_slot(cfg, s) and assignment.get(s.slot_id)
    ]
    if len(heavy_slots) < 2:
        return assignment, []

    candidates_by_slot = {
        s.slot_id: [norm_name(d) for d in dict.fromkeys(s.allowed or []) if norm_name(d) and norm_name(d) != "Recupero"]
        for s in heavy_slots
    }
    candidate_docs = sorted({d for vals in candidates_by_slot.values() for d in vals})
    if len(candidate_docs) < 2:
        return assignment, []

    slots_by_date: Dict[dt.date, List[Slot]] = defaultdict(list)
    j_slots = []
    for s in slots:
        slots_by_date[s.day.date].append(s)
        if s.columns == ["J"]:
            j_slots.append(s)

    gc = cfg.get("global_constraints") or {}
    night_gap = int(gc.get("night_spacing_days_min", 5) or 5)
    night_off = gc.get("night_off") if isinstance(gc.get("night_off"), dict) else {}
    night_off_next_day = bool(night_off.get("next_day", True))

    def assigned_on_date(assign: Dict[str, Optional[str]], doc: str, day: dt.date, exclude_sid: Optional[str] = None) -> bool:
        for ds in slots_by_date.get(day, []):
            if ds.slot_id == exclude_sid:
                continue
            if ds.columns == ["C"]:
                continue
            if norm_name(assign.get(ds.slot_id)) == doc:
                return True
        return False

    def has_j_on(assign: Dict[str, Optional[str]], doc: str, day: dt.date, exclude_sid: Optional[str] = None) -> bool:
        for js in j_slots:
            if js.slot_id == exclude_sid:
                continue
            if js.day.date == day and norm_name(assign.get(js.slot_id)) == doc:
                return True
        return False

    def safe_replacement(s: Slot, doc: str, assign: Dict[str, Optional[str]]) -> bool:
        if s.slot_id in fixed_ids:
            return False
        if doc not in candidates_by_slot.get(s.slot_id, []):
            return False
        if assigned_on_date(assign, doc, s.day.date, exclude_sid=s.slot_id):
            return False
        if s.columns == ["J"]:
            for js in j_slots:
                if js.slot_id == s.slot_id:
                    continue
                if norm_name(assign.get(js.slot_id)) != doc:
                    continue
                delta = abs((js.day.date - s.day.date).days)
                if 0 < delta < night_gap:
                    return False
            if night_off_next_day and assigned_on_date(assign, doc, s.day.date + dt.timedelta(days=1)):
                return False
        else:
            if has_j_on(assign, doc, s.day.date):
                return False
            if night_off_next_day and has_j_on(assign, doc, s.day.date - dt.timedelta(days=1)):
                return False
        return True

    heavy_dates = sorted({s.day.date for s in heavy_slots})

    def score(assign: Dict[str, Optional[str]]) -> Tuple[int, int, int, int, int, int]:
        load = Counter()
        by_doc_dates: Dict[str, Set[dt.date]] = defaultdict(set)
        for hs in heavy_slots:
            doc = norm_name(assign.get(hs.slot_id))
            if not doc or doc == "Recupero":
                continue
            if doc not in candidate_docs:
                continue
            load[doc] += 1
            by_doc_dates[doc].add(hs.day.date)
        zero_candidates = sum(1 for doc in candidate_docs if load.get(doc, 0) == 0)
        max_load = max((load.get(doc, 0) for doc in candidate_docs), default=0)
        duplicate_load = sum(max(0, load.get(doc, 0) - 1) for doc in candidate_docs)
        sum_squares = sum(load.get(doc, 0) ** 2 for doc in candidate_docs)
        consecutive = 0
        for doc in candidate_docs:
            dates = by_doc_dates.get(doc, set())
            for a, b in zip(heavy_dates, heavy_dates[1:]):
                if a in dates and b in dates:
                    consecutive += 1
        repeated_same_date = 0
        for doc in candidate_docs:
            per_date = Counter(
                hs.day.date
                for hs in heavy_slots
                if norm_name(assign.get(hs.slot_id)) == doc
            )
            repeated_same_date += sum(max(0, n - 1) for n in per_date.values())
        return (zero_candidates, max_load, duplicate_load, consecutive, repeated_same_date, sum_squares)

    repaired = dict(assignment)
    swaps: List[Dict] = []
    current_score = score(repaired)

    for _ in range(max_passes):
        best = None
        best_score = current_score
        for s in sorted(heavy_slots, key=lambda x: (x.day.date, x.slot_id)):
            if s.slot_id in fixed_ids:
                continue
            current_doc = norm_name(repaired.get(s.slot_id))
            if not current_doc:
                continue
            for cand in candidates_by_slot.get(s.slot_id, []):
                if cand == current_doc:
                    continue
                if not safe_replacement(s, cand, repaired):
                    continue
                trial = dict(repaired)
                trial[s.slot_id] = cand
                trial_score = score(trial)
                if trial_score < best_score:
                    best_score = trial_score
                    best = (s, current_doc, cand, trial, trial_score)
        if best is None:
            break
        s, old_doc, new_doc, repaired, new_score = best
        swaps.append({
            "date": s.day.date.isoformat(),
            "slot_id": s.slot_id,
            "columns": list(s.columns),
            "from": old_doc,
            "to": new_doc,
            "score_before": current_score,
            "score_after": new_score,
        })
        current_score = new_score

    return repaired, swaps


def solve_with_ortools(
    cfg: dict,
    days: List[DayRow],
    slots: List[Slot],
    fixed_assignments: Optional[List[dict]] = None,
    availability_preferences: Optional[List[dict]] = None,
    unav_map: Optional[Dict[str, Dict[dt.date, Set[str]]]] = None,
    historical_stats: Optional[dict] = None,
    prior_usage: Optional[dict] = None,
) -> Tuple[Dict[str, Optional[str]], Dict]:
    """
    Returns:
      assignment: slot_id -> doctor (or None for optional left blank)
      stats: dict with diagnostic info

    fixed_assignments: [{"doctor": str, "date": "YYYY-MM-DD", "column": str}, ...]
      Vincolo HARD: quel medico DEVE comparire in quella colonna quel giorno.
    availability_preferences: [{"doctor": str, "date": "YYYY-MM-DD", "shift": str}, ...]
      Vincolo SOFT: il solver prova a far comparire il medico in qualsiasi slot
      della fascia indicata in quel giorno.
    unav_map: mappa indisponibilità, usata per calcolare il target universitari corretto.
    """
    try:
        from ortools.sat.python import cp_model
    except Exception as e:
        raise RuntimeError("OR-Tools not installed. Install with: pip install ortools") from e
    model = cp_model.CpModel()
    # Collect extra objective terms built during constraint setup
    extra_obj = []
    heavy_priority_terms = []
    # Warnings raccolti durante la costruzione del modello (non bloccanti)
    pre_solve_warnings: List[str] = []

    # Diagnostics for the "Recupero su T in 2 lunedì" rule.
    # NOTE: this MUST be defined in the OR-Tools path too; otherwise the code may
    # crash after solving and trigger the greedy fallback.
    MonT_need_rec = 0
    MonT_target_dates: List[dt.date] = []
    doctors = collect_doctors(cfg)
    doctors = [d for d in doctors if d != "Recupero"] + (["Recupero"] if "Recupero" in doctors else [])
    doc_to_idx = {d:i for i,d in enumerate(doctors)}
    month_key_for_solver = f"{days[0].date.year:04d}-{days[0].date.month:02d}" if days else ""
    prior_usage = prior_usage or {}
    prior_counts_month = ((prior_usage.get("counts") or {}).get(month_key_for_solver) or {})

    period_month_fraction = 1.0
    period_is_partial_month = False
    if days:
        try:
            _period_dates = sorted({d.date for d in days})
            _first_period_day = _period_dates[0]
            _last_period_day = _period_dates[-1]
            _full_month_days = calendar.monthrange(_first_period_day.year, _first_period_day.month)[1]
            period_is_partial_month = (
                _first_period_day.day != 1
                or _last_period_day.day != _full_month_days
                or len(_period_dates) != _full_month_days
            )
            if period_is_partial_month and _full_month_days > 0:
                period_month_fraction = len(_period_dates) / _full_month_days
                pre_solve_warnings.append(
                    "Quote mensili riproporzionate al periodo: "
                    f"{len(_period_dates)}/{_full_month_days} giorni del mese."
                )
        except Exception:
            period_month_fraction = 1.0
            period_is_partial_month = False

    def _date_is_j_festive(d: dt.date) -> bool:
        if d.weekday() in (5, 6):
            return True
        try:
            if d in italy_public_holidays(int(d.year)):
                return True
        except Exception:
            pass
        for raw in cfg.get("festivi_extra", []) or []:
            try:
                if parse_date(raw) == d:
                    return True
            except Exception:
                continue
        return False

    def _day_is_j_festive(day: DayRow) -> bool:
        return _date_is_j_festive(day.date)

    def _prior_count(doc: Optional[str], col: Optional[str]) -> int:
        if not doc or not col:
            return 0
        doc_n = norm_name(doc)
        col_n = str(col).strip().upper()
        by_col = prior_counts_month.get(doc_n) or prior_counts_month.get(str(doc).strip()) or {}
        period_by_col = ((prior_usage.get("period_counts") or {}).get(doc_n)
                         or (prior_usage.get("period_counts") or {}).get(str(doc).strip())
                         or {})
        try:
            if col_n == "FESTIVI":
                direct = int(by_col.get("Festivi", by_col.get("FESTIVI", 0)) or 0)
                if direct <= 0:
                    direct = sum(int(by_col.get(c, 0) or 0) for c in ("D", "E", "H", "I"))
                return direct + int(period_by_col.get("Festivi", period_by_col.get("FESTIVI", 0)) or 0)
            return int(by_col.get(col_n, 0) or 0) + int(period_by_col.get(col_n, 0) or 0)
        except Exception:
            return 0

    def _prior_total_count(doc: Optional[str]) -> int:
        if not doc:
            return 0
        doc_n = norm_name(doc)
        by_col = prior_counts_month.get(doc_n) or prior_counts_month.get(str(doc).strip()) or {}
        total = 0
        for v in by_col.values():
            try:
                total += int(v or 0)
            except Exception:
                continue
        return total

    def _prior_j_weekend_count(doc: Optional[str]) -> int:
        if not doc:
            return 0
        doc_n = norm_name(doc)
        period_by_col = ((prior_usage.get("period_counts") or {}).get(doc_n)
                         or (prior_usage.get("period_counts") or {}).get(str(doc).strip())
                         or {})
        by_doc = ((prior_usage.get("night_dates_by_doc") or {}).get(month_key_for_solver) or {})
        dates = by_doc.get(doc_n) or by_doc.get(str(doc).strip()) or []
        total = int(period_by_col.get("J_FESTIVI", 0) or 0)
        for ds in dates:
            try:
                d = dt.date.fromisoformat(str(ds)[:10])
            except Exception:
                continue
            if _date_is_j_festive(d):
                total += 1
        return total

    def _prior_j_weekday_count(doc: Optional[str]) -> int:
        if not doc:
            return 0
        doc_n = norm_name(doc)
        by_doc = ((prior_usage.get("night_dates_by_doc") or {}).get(month_key_for_solver) or {})
        dates = by_doc.get(doc_n) or by_doc.get(str(doc).strip()) or []
        total = 0
        for ds in dates:
            try:
                d = dt.date.fromisoformat(str(ds)[:10])
            except Exception:
                continue
            if not _date_is_j_festive(d):
                total += 1
        return total

    def _prior_j_sunday_count(doc: Optional[str]) -> int:
        if not doc:
            return 0
        doc_n = norm_name(doc)
        by_doc = ((prior_usage.get("night_dates_by_doc") or {}).get(month_key_for_solver) or {})
        dates = by_doc.get(doc_n) or by_doc.get(str(doc).strip()) or []
        total = 0
        for ds in dates:
            try:
                d = dt.date.fromisoformat(str(ds)[:10])
            except Exception:
                continue
            if d.weekday() == 6:
                total += 1
        return total

    def _prior_rule_balance_count(doc: Optional[str], rule_key: Optional[str]) -> int:
        if not doc or not rule_key:
            return 0
        key = str(rule_key).strip().upper()
        pair_map = {
            "E_G": ("E", "G"),
            "FESTIVO_DE": ("D", "E"),
            "FESTIVO_HI": ("H", "I"),
            "DE": ("D", "E"),
            "HI": ("H", "I"),
        }
        cols = pair_map.get(key)
        if cols is None and "_" in key:
            cols = tuple(c for c in key.split("_") if c)
        if cols:
            return max((_prior_count(doc, c) for c in cols), default=0)
        return _prior_count(doc, key)

    def _has_prior_month_usage() -> bool:
        return bool(prior_counts_month)

    def _monthly_target_for_period(value: int, doc: Optional[str] = None, col: Optional[str] = None) -> int:
        value = int(value)
        if period_is_partial_month and doc and col and _has_prior_month_usage():
            return max(0, value - _prior_count(doc, col))
        if not period_is_partial_month:
            return value
        return int(math.floor(value * period_month_fraction + 0.5))

    def _monthly_cap_for_period(value: int, doc: Optional[str] = None, col: Optional[str] = None) -> int:
        value = int(value)
        if period_is_partial_month and doc and col and _has_prior_month_usage():
            return max(0, value - _prior_count(doc, col))
        if not period_is_partial_month:
            return value
        return max(1, int(math.ceil(value * period_month_fraction))) if value > 0 else 0

    # Fixed assignments are admin overrides, but they must be part of the slot
    # domain before variables and slot cardinality constraints are created.
    for fa in (fixed_assignments or []):
        try:
            fa_date = dt.date.fromisoformat(str(fa.get("date","")).strip())
            fa_doc = norm_name(str(fa.get("doctor","")).strip())
            fa_col = str(fa.get("column","")).strip().upper()
        except Exception:
            continue
        if fa_doc not in doc_to_idx:
            pre_solve_warnings.append(
                f"Fixed assignment ignorata: medico sconosciuto {fa_doc} su {fa_date}/{fa_col}."
            )
            continue
        ok_fixed, reason = _fixed_assignment_allowed(cfg, fa_col, fa_doc)
        if not ok_fixed:
            pre_solve_warnings.append(
                f"Fixed assignment ignorata: {fa_doc} su {fa_date}/{fa_col} ({reason})."
            )
            continue
        for s in slots:
            if s.day.date == fa_date and fa_col in [str(c).strip().upper() for c in (s.columns or [])]:
                if fa_doc not in s.allowed:
                    s.allowed.append(fa_doc)
                    existing = s.emergency_doctors or []
                    if fa_doc not in existing:
                        s.emergency_doctors = existing + [fa_doc]

    # Decision vars: x[(slot_id, doc)] in {0,1}
    x = {}
    for s in slots:
        for d in s.allowed:
            if d not in doc_to_idx:
                continue
            x[(s.slot_id, d)] = model.NewBoolVar(f"x_{hash(s.slot_id)%10**8}_{hash(d)%10**8}")
    # Slot assignment constraints.
    # Required slots use a "soft-required" approach: sum(vars) + b_blank == 1.
    # b_blank == 1 means the slot is left empty (penalized by severity tier).
    # This makes the model ALWAYS feasible and lets the solver identify
    # which slots genuinely cannot be filled (they show up blank in the output).
    #
    # Gerarchia sacrifici (crescente = si sacrifica per primo):
    #   K=T share/preferenze/quote/target  →  slot obbligatorio vuoto.
    # Il vuoto obbligatorio deve restare una vera ultima spiaggia: in particolare
    # una J copribile non deve mai saltare per "migliorare" un bilanciamento.
    BLANK_REQUIRED_PENALTY = 500_000_000   # default per slot non classificati
    _BLANK_PENALTY_BY_TAG: Dict[str, int] = {
        # Critici — quasi mai vuoti
        "J":        2_000_000_000,
        "K":        1_500_000_000,
        "H":        1_500_000_000,
        "D_F.D":    1_500_000_000,
        "D_F.F":    1_500_000_000,
        "Festivo_HI": 1_500_000_000,
        # Alti — solo se davvero nessun medico disponibile
        "I":        1_000_000_000,
        "E_G":      1_000_000_000,
        "Festivo_DE": 1_000_000_000,
        "T":          800_000_000,
        "Q":          800_000_000,
        "V":          800_000_000,
        # Bassi — possono cedere prima degli altri required
        "W":          300_000_000,
    }
    blank_required_vars: Dict[str, object] = {}  # slot_id -> b_blank var (for diagnostics)
    coverage_obj_terms: List = []
    for s in slots:
        vars_ = [x[(s.slot_id, d)] for d in s.allowed if (s.slot_id, d) in x]
        if not vars_:
            # No eligible doctors at all → slot will be blank (no variable to add penalty to)
            continue
        if s.required:
            b_blank = model.NewBoolVar(f"blank_req_{hash(s.slot_id)%10**8}")
            model.Add(sum(vars_) + b_blank == 1)
            penalty = _BLANK_PENALTY_BY_TAG.get(s.rule_tag, BLANK_REQUIRED_PENALTY)
            coverage_obj_terms.append(penalty * b_blank)
            blank_required_vars[s.slot_id] = b_blank
        else:
            # Optional slot: can be left blank. If blank_penalty>0, we penalize blanks so it is used only as a last resort.
            if getattr(s, "blank_penalty", 0) and int(getattr(s, "blank_penalty", 0)) > 0:
                # Optional slot may be left blank, but only as a last resort: we add a large penalty
                # term to the objective. This must go into `extra_obj` because `objective_terms`
                # is defined later, when all high-level soft constraints are assembled.
                b = model.NewBoolVar(f"blank_{hash(s.slot_id)%10**8}")
                model.Add(sum(vars_) + b == 1)
                extra_obj.append(b * int(getattr(s, "blank_penalty", 0)))
            else:
                model.Add(sum(vars_) <= 1)
    # Penalità per medici di emergenza (non nel pool primario della colonna)
    # Il solver li usa solo se non c'è alternativa migliore (penalità < blank penalty).
    EMERGENCY_FILL_PENALTY = 5_000_000  # < BLANK_REQUIRED_PENALTY → meglio di blank
    for s in slots:
        for _emerg_doc in (s.emergency_doctors or []):
            _ev = x.get((s.slot_id, norm_name(_emerg_doc)))
            if _ev is not None:
                extra_obj.append(EMERGENCY_FILL_PENALTY * _ev)

    # Helper: slots per day
    slots_by_day: Dict[dt.date, List[Slot]] = defaultdict(list)
    for s in slots:
        slots_by_day[s.day.date].append(s)

    # ---------------------------------------------------------------
    # ASSEGNAZIONI FISSE (hard): l'admin ha fissato un medico in una
    # colonna specifica in un giorno specifico.
    # ---------------------------------------------------------------
    for fa in (fixed_assignments or []):
        try:
            fa_date = dt.date.fromisoformat(str(fa.get("date","")).strip())
            fa_doc = norm_name(str(fa.get("doctor","")).strip())
            fa_col = str(fa.get("column","")).strip().upper()
        except Exception:
            continue
        if fa_doc not in doc_to_idx:
            continue
        # Trova lo slot corrispondente
        target_slots = [s for s in slots_by_day.get(fa_date, [])
                        if fa_col in [str(c).strip().upper() for c in (s.columns or [])]]
        if not target_slots:
            continue
        for ts in target_slots:
            tv = x.get((ts.slot_id, fa_doc))
            if tv is None:
                pre_solve_warnings.append(
                    f"Fixed assignment ignorata: {fa_doc} non ammesso in {ts.slot_id}."
                )
                continue
            model.Add(tv == 1)  # HARD: questo medico deve essere assegnato
            # Forza tutti gli altri a 0 in questo slot
            for d2 in list(ts.allowed) + list(doctors):
                d2n = norm_name(d2)
                if d2n == fa_doc:
                    continue
                v2 = x.get((ts.slot_id, d2n))
                if v2 is not None:
                    model.Add(v2 == 0)

    # ---------------------------------------------------------------
    # DISPONIBILITÀ (soft): il medico VUOLE comparire in una certa fascia.
    # Penalizziamo se NON compare in nessuno slot di quella fascia in quel giorno.
    # ---------------------------------------------------------------
    SHIFT_MAP_AVAIL = {
        "mattina": "Mattina", "morning": "Mattina",
        "pomeriggio": "Pomeriggio", "afternoon": "Pomeriggio",
        "notte": "Notte", "night": "Notte",
        "diurno": "Diurno", "day": "Diurno",
    }
    # Penalità per disponibilità non rispettata.
    # Calibrazione rispetto alle penalità strutturali D/F:
    #   missing_pair_doc=12000, df_share=8000, prefer_df_share=15000
    #   → stack peggiore ~35000
    # "media" (40000) supera ogni singola penalità strutturale e lo stack comune.
    # "alta"  (80000) supera anche gli stack peggiori.
    # "bassa" (15000) può essere superata da singole penalità strutturali (intenzionale).
    AVAIL_PENALTY_BASE = 40_000
    AVAIL_PRIORITY_MULT = {"alta": 2.0, "alta priority": 2.0, "media": 1.0, "bassa": 0.375}
    pref_skipped_log: list = []
    for ap in (availability_preferences or []):
        try:
            ap_date = dt.date.fromisoformat(str(ap.get("date","")).strip())
            ap_doc = norm_name(str(ap.get("doctor","")).strip())
            ap_shift_raw = str(ap.get("shift","")).strip().lower()
            ap_shift = SHIFT_MAP_AVAIL.get(ap_shift_raw, ap_shift_raw.capitalize())
            ap_priority = str(ap.get("priority", "media")).strip().lower()
        except Exception:
            continue
        if ap_doc not in doc_to_idx:
            continue
        mult = AVAIL_PRIORITY_MULT.get(ap_priority, 1.0)
        penalty = int(AVAIL_PENALTY_BASE * mult)
        # Raccoglie tutti i vars dello slot per quella fascia in quel giorno
        ap_vars = []
        for s in slots_by_day.get(ap_date, []):
            s_shift = getattr(s, "shift", "") or ""
            if ap_shift.lower() in s_shift.lower() or s_shift.lower() in ap_shift.lower():
                v = x.get((s.slot_id, ap_doc))
                if v is not None:
                    ap_vars.append(v)
        if ap_vars:
            # b = 1 se il medico compare in almeno uno slot di quella fascia
            b_avail = model.NewBoolVar(f"avail_{ap_date}_{hash(ap_doc)%10**6}_{ap_shift}")
            model.AddMaxEquality(b_avail, ap_vars)
            not_avail = model.NewBoolVar(f"notavail_{ap_date}_{hash(ap_doc)%10**6}_{ap_shift}")
            model.Add(not_avail + b_avail == 1)
            extra_obj.append(penalty * not_avail)
        else:
            pref_skipped_log.append(
                f"  AVAIL SKIP {ap_doc} {ap_date} {ap_shift}: non in nessuno slot ammesso"
            )
    if pref_skipped_log:
        print("WARNING: Preferenze disponibilità non applicabili (medico fuori pool):\n" +
              "\n".join(pref_skipped_log))


    # Uniqueness per day: one doctor max 1 slot/day (exceptions already handled by merged columns)
    gc = cfg.get("global_constraints") or {}
    relief = gc.get("relief_valves") or {}
    service_combos = cfg.get("pool_service_combinations") or []
    kt_combo_fallback = any(
        tuple(sorted(str(c).strip().upper() for c in (combo.get("columns") or []))) == ("K", "T")
        and str(combo.get("mode", "")).strip().lower() in {"fallback", "always"}
        for combo in service_combos
        if isinstance(combo, dict)
    )
    enable_kt_share = bool(relief.get("enable_kt_share", False)) or kt_combo_fallback
    # K/T is a rescue valve only: keep the penalty above ordinary soft costs, but
    # below the required-slot blank penalty for K/T so it can still save coverage.
    KT_SHARE_PENALTY_FLOOR = 4_000_000
    kt_share_penalty = max(int(relief.get("kt_share_penalty", 5000)), KT_SHARE_PENALTY_FLOOR)

    # D/F share valve (allows the same doctor to cover both D and F ONLY if needed)
    rules_map = cfg.get("rules", {}) or {}
    r_df = rules_map.get("D_F", {}) if isinstance(rules_map.get("D_F", {}), dict) else {}
    enable_df_share = bool(r_df.get("enable_df_share", True))
    df_share_penalty = int(r_df.get("df_share_penalty", 8000))
    prefer_df_share_penalty = int(r_df.get("prefer_df_share_penalty", 15000))
    df_share_vars_by_date: Dict[dt.date, Dict[str, object]] = {}

    daily_exempt_cols = {str(c).strip().upper() for c in (gc.get('daily_uniqueness_exempt_columns') or [])}
    def _slot_is_exempt_daily(s: Slot) -> bool:
        return any(str(c).strip().upper() in daily_exempt_cols for c in (s.columns or []))

    for day in days:
        day_slots = slots_by_day.get(day.date, [])
        uniq_slots = [s for s in day_slots if not _slot_is_exempt_daily(s)]

        # Slack vars per doctor (can allow +1 assignment for a specific "share" valve)
        slack_vars_by_doc: Dict[str, List] = defaultdict(list)

        # ---- K/T emergency share
        kt_y_by_doc = None
        slotK = None
        slotT = None
        if enable_kt_share:
            for s in day_slots:
                if s.columns == ["K"]:
                    slotK = s
                elif s.columns == ["T"]:
                    slotT = s
        if enable_kt_share and slotK is not None and slotT is not None:
            kt_y_by_doc = {}
            for d in doctors:
                if (slotK.slot_id, d) in x and (slotT.slot_id, d) in x:
                    y = model.NewBoolVar(f"kt_same_{day.date.isoformat()}_{hash(d)%10**6}")
                    model.Add(y <= x[(slotK.slot_id, d)])
                    model.Add(y <= x[(slotT.slot_id, d)])
                    model.Add(y >= x[(slotK.slot_id, d)] + x[(slotT.slot_id, d)] - 1)
                    kt_y_by_doc[d] = y
            if kt_y_by_doc:
                model.Add(sum(kt_y_by_doc.values()) <= 1)
                extra_obj.append(sum(kt_y_by_doc.values()) * kt_share_penalty)
                for d, y in kt_y_by_doc.items():
                    slack_vars_by_doc[d].append(y)

        # ---- D/F emergency share (used only when the D/F pair is infeasible without it)
        df_y_by_doc = None
        slotD = None
        slotF = None
        if enable_df_share:
            for s in day_slots:
                if s.columns == ["D"]:
                    slotD = s
                elif s.columns == ["F"]:
                    slotF = s
        if enable_df_share and slotD is not None and slotF is not None:
            df_y_by_doc = {}
            for d in doctors:
                if (slotD.slot_id, d) in x and (slotF.slot_id, d) in x:
                    y = model.NewBoolVar(f"df_same_{day.date.isoformat()}_{hash(d)%10**6}")
                    model.Add(y <= x[(slotD.slot_id, d)])
                    model.Add(y <= x[(slotF.slot_id, d)])
                    model.Add(y >= x[(slotD.slot_id, d)] + x[(slotF.slot_id, d)] - 1)
                    df_y_by_doc[d] = y
            if df_y_by_doc:
                df_share_vars_by_date[day.date] = df_y_by_doc
                model.Add(sum(df_y_by_doc.values()) <= 1)

                # If D/F is in an "emergency" mode (both Grimaldi/Calabrò unavailable, or H pool empty with only one available),
                # we PREFER using the same doctor (D=F) to avoid blanks/infeasible. In that case we do NOT penalize df_share,
                # and we penalize NOT sharing.
                prefer_share = bool(getattr(slotD, "force_same_doctor", False)) and bool(getattr(slotF, "force_same_doctor", False))
                share_pen = 0 if prefer_share else df_share_penalty
                if share_pen:
                    extra_obj.append(sum(df_y_by_doc.values()) * share_pen)

                if prefer_share:
                    share_any = model.NewBoolVar(f"df_share_any_{day.date.isoformat()}")
                    model.AddMaxEquality(share_any, list(df_y_by_doc.values()))
                    notshare = model.NewBoolVar(f"df_notshare_{day.date.isoformat()}")
                    model.Add(notshare + share_any == 1)
                    extra_obj.append(notshare * prefer_df_share_penalty)

                for d, y in df_y_by_doc.items():
                    slack_vars_by_doc[d].append(y)

        # Prevent a single doctor from stacking multiple share valves on the same day (safety)
        if kt_y_by_doc and df_y_by_doc:
            for d in doctors:
                y1 = kt_y_by_doc.get(d)
                y2 = df_y_by_doc.get(d)
                if y1 is not None and y2 is not None:
                    model.Add(y1 + y2 <= 1)

        # Daily uniqueness constraint with optional +1 slack per share valve
        for d in doctors:
            vars_ = []
            for s in uniq_slots:
                if (s.slot_id, d) in x:
                    vars_.append(x[(s.slot_id, d)])
            if not vars_:
                continue
            slack_terms = slack_vars_by_doc.get(d, [])
            if slack_terms:
                model.Add(sum(vars_) <= 1 + sum(slack_terms))
            else:
                model.Add(sum(vars_) <= 1)
# ---- H/I: divieto di stesso medico in giorni consecutivi (H e I indipendenti)
    # Vale anche sui Festivi, dove esiste lo slot unico HI (colonne H+I).
    try:
        slot_ids_by_date_col: Dict[Tuple[dt.date, str], List[str]] = defaultdict(list)
        for s in slots:
            for c in (s.columns or []):
                cc = str(c).strip().upper()
                if cc in {"H", "I"}:
                    slot_ids_by_date_col[(s.day.date, cc)].append(s.slot_id)
        days_sorted = sorted(days, key=lambda d: d.date)
        for i in range(len(days_sorted) - 1):
            d1 = days_sorted[i].date
            d2 = days_sorted[i + 1].date
            for col in ("H", "I"):
                sids1 = slot_ids_by_date_col.get((d1, col), [])
                sids2 = slot_ids_by_date_col.get((d2, col), [])
                if not sids1 and not sids2:
                    continue
                for doc in doctors:
                    if doc == "Recupero":
                        continue
                    v1 = [x.get((sid, doc)) for sid in sids1 if (sid, doc) in x]
                    v2 = [x.get((sid, doc)) for sid in sids2 if (sid, doc) in x]
                    v1 = [v for v in v1 if v is not None]
                    v2 = [v for v in v2 if v is not None]
                    if v1 or v2:
                        model.Add(sum(v1) + sum(v2) <= 1)
    except Exception:
        pass

    # H is a flexible afternoon slot, but several H doctors are also the scarce
    # morning pool for K/L/T/Q/R/E+G/etc. Prefer using the H candidate with fewer
    # same-day morning alternatives (e.g. Migliorato) so multi-service doctors
    # remain available for bottleneck columns.
    try:
        rH_pref = rules_map.get("H", {}) if isinstance(rules_map.get("H", {}), dict) else {}
        h_reserve_penalty = int(rH_pref.get("reserve_multi_service_penalty", 1_000_000) or 0)
        h_reserve_cols_raw = rH_pref.get("reserve_for_columns")
        if h_reserve_cols_raw is None:
            h_reserve_cols = {"E", "G", "K", "L", "Q", "R", "S", "T", "U", "V", "W", "Y", "Z", "AB"}
        else:
            h_reserve_cols = {
                str(c).strip().upper()
                for c in h_reserve_cols_raw
                if str(c).strip()
            }
        if h_reserve_penalty > 0 and h_reserve_cols:
            for day in days:
                h_slot = next((s for s in slots_by_day.get(day.date, []) if s.columns == ["H"]), None)
                if not h_slot:
                    continue
                pressure_slots = [
                    s for s in slots_by_day.get(day.date, [])
                    if s.slot_id != h_slot.slot_id
                    and getattr(s, "shift", None) == "Mattina"
                    and any(str(c).strip().upper() in h_reserve_cols for c in (s.columns or []))
                ]
                if not pressure_slots:
                    continue
                scores: Dict[str, int] = {}
                for doc in h_slot.allowed:
                    hv = x.get((h_slot.slot_id, doc))
                    if hv is None:
                        continue
                    scores[doc] = sum(
                        1
                        for ps in pressure_slots
                        if x.get((ps.slot_id, doc)) is not None
                    )
                if len(scores) < 2:
                    continue
                min_score = min(scores.values())
                max_score = max(scores.values())
                if max_score <= min_score:
                    continue
                for doc, score in scores.items():
                    if score <= min_score:
                        continue
                    hv = x.get((h_slot.slot_id, doc))
                    if hv is not None:
                        extra_obj.append(h_reserve_penalty * (score - min_score) * hv)
    except Exception:
        pass

    # K is often the bottleneck during heavy vacation periods. When the same-day
    # K pool is already narrow, avoid spending K-capable doctors on flexible
    # H/Q/T-style services if that service has an alternative outside the K pool.
    try:
        k_preserve_penalty = int(gc.get("k_pool_preserve_penalty", 80_000_000) or 0)
        k_preserve_threshold = int(gc.get("k_pool_preserve_threshold", 5) or 0)
        k_pressure_cols = {
            "H", "L", "Q", "R", "S", "T", "U", "V", "W", "Y", "Z", "AB"
        }
        if k_preserve_penalty > 0:
            for day in days:
                k_slot = next((s for s in slots_by_day.get(day.date, []) if s.columns == ["K"]), None)
                if not k_slot or not k_slot.allowed:
                    continue
                k_docs = {norm_name(d) for d in k_slot.allowed if x.get((k_slot.slot_id, norm_name(d))) is not None}
                if not k_docs:
                    continue
                if k_preserve_threshold > 0 and len(k_docs) > k_preserve_threshold:
                    continue
                for other in slots_by_day.get(day.date, []):
                    if other.slot_id == k_slot.slot_id:
                        continue
                    if any(str(c).strip().upper() in {"C", "J", "D", "E", "F", "G", "I"} for c in (other.columns or [])):
                        continue
                    if not any(str(c).strip().upper() in k_pressure_cols for c in (other.columns or [])):
                        continue
                    other_docs = {norm_name(d) for d in other.allowed if x.get((other.slot_id, norm_name(d))) is not None}
                    if not (other_docs & k_docs):
                        continue
                    if not (other_docs - k_docs):
                        continue
                    for doc in sorted(other_docs & k_docs):
                        ov = x.get((other.slot_id, doc))
                        if ov is not None:
                            extra_obj.append(k_preserve_penalty * ov)
    except Exception:
        pass

    # Protect very constrained morning services such as E+G. If a doctor is one
    # of the few candidates for E+G, avoid spending that doctor on another
    # morning service that has alternatives outside the E+G pool.
    try:
        scarce_penalty = int(((rules_map.get("E_G", {}) or {}).get("protect_pool_penalty", 35_000_000)) or 0)
        protected_other_cols = {"K", "L", "Q", "R", "S", "T", "U", "V", "W", "Y", "Z"}
        if scarce_penalty > 0:
            for day in days:
                eg_slot = next((s for s in slots_by_day.get(day.date, []) if s.columns == ["E", "G"]), None)
                if not eg_slot or not eg_slot.allowed:
                    continue
                protected_docs = set(eg_slot.allowed)
                for other in slots_by_day.get(day.date, []):
                    if other.slot_id == eg_slot.slot_id:
                        continue
                    if getattr(other, "shift", None) != "Mattina":
                        continue
                    if not any(str(c).strip().upper() in protected_other_cols for c in (other.columns or [])):
                        continue
                    other_has_outside_choice = any(
                        doc not in protected_docs and x.get((other.slot_id, doc)) is not None
                        for doc in other.allowed
                    )
                    if not other_has_outside_choice:
                        continue
                    for doc in protected_docs:
                        ov = x.get((other.slot_id, doc))
                        if ov is not None:
                            extra_obj.append(scarce_penalty * ov)
    except Exception:
        pass

    # Same idea for tiny pools such as I and AB. If only Allegra/Crea can cover
    # I/AB on a day, do not spend them on K/T/Q/R/etc. when those services have
    # alternatives outside that tiny pool.
    try:
        tiny_pool_penalty = int((gc.get("tiny_pool_protect_penalty", 45_000_000) or 45_000_000))
        tiny_pool_tags = {"I", "AB"}
        tiny_other_cols = {"C", "K", "L", "Q", "R", "S", "T", "U", "V", "W", "Y", "Z"}
        if tiny_pool_penalty > 0:
            for day in days:
                scarce_slots = [
                    s for s in slots_by_day.get(day.date, [])
                    if s.required
                    and s.rule_tag in tiny_pool_tags
                    and 1 <= len(s.allowed or []) <= 3
                ]
                if not scarce_slots:
                    continue
                scarce_slot_ids = {s.slot_id for s in scarce_slots}
                protected_docs = {doc for s in scarce_slots for doc in (s.allowed or [])}
                if not protected_docs:
                    continue
                for other in slots_by_day.get(day.date, []):
                    if other.slot_id in scarce_slot_ids:
                        continue
                    if not any(str(c).strip().upper() in tiny_other_cols for c in (other.columns or [])):
                        continue
                    other_has_outside_choice = any(
                        doc not in protected_docs and x.get((other.slot_id, doc)) is not None
                        for doc in other.allowed
                    )
                    if not other_has_outside_choice:
                        continue
                    for doc in protected_docs:
                        ov = x.get((other.slot_id, doc))
                        if ov is not None:
                            extra_obj.append(tiny_pool_penalty * ov)
    except Exception:
        pass

    # Also protect tiny next-day pools from night assignments: if Allegra/Crea
    # are the only doctors for I/AB tomorrow, avoid assigning them to J tonight
    # when the night pool has alternatives, because night_off_next_day would
    # remove them from tomorrow's scarce service.
    try:
        next_day_tiny_penalty = int((gc.get("next_day_tiny_pool_night_penalty", 60_000_000) or 60_000_000))
        if next_day_tiny_penalty > 0:
            day_by_date = {d.date: d for d in days}
            for day in days:
                tomorrow = day.date + dt.timedelta(days=1)
                if tomorrow not in day_by_date:
                    continue
                j_slot = next((s for s in slots_by_day.get(day.date, []) if s.columns == ["J"]), None)
                if not j_slot:
                    continue
                scarce_slots = [
                    s for s in slots_by_day.get(tomorrow, [])
                    if s.required
                    and s.rule_tag in {"I", "AB"}
                    and 1 <= len(s.allowed or []) <= 3
                ]
                if not scarce_slots:
                    continue
                protected_docs = {doc for s in scarce_slots for doc in (s.allowed or [])}
                if not protected_docs:
                    continue
                j_has_outside_choice = any(
                    doc not in protected_docs and x.get((j_slot.slot_id, doc)) is not None
                    for doc in j_slot.allowed
                )
                if not j_has_outside_choice:
                    continue
                for doc in protected_docs:
                    jv = x.get((j_slot.slot_id, doc))
                    if jv is not None:
                        extra_obj.append(next_day_tiny_penalty * jv)
    except Exception:
        pass

    # ---- D/F pattern 3+3 for Grimaldi/Calabrò (conditional hard)
    # If configured in rules.D_F (pattern_3_3 + pattern_conditional_hard),
    # enforce:
    #   Mon–Wed: D=doc1, F=doc2
    #   Thu–Sat: D=doc2, F=doc1
    # Pattern can be violated ONLY on days where (H is doc1/doc2) OR (J is doc2).
    rules_map = cfg.get("rules", {}) or {}
    r_df = rules_map.get("D_F", {}) if isinstance(rules_map.get("D_F", {}), dict) else {}
    if r_df.get("pattern_3_3") and r_df.get("pattern_conditional_hard"):
        doc1 = norm_name(r_df.get("pattern_doc1") or "")
        doc2 = norm_name(r_df.get("pattern_doc2") or "")
        if doc1 in doctors and doc2 in doctors:
            _day_idx_map = {d.date: i for i, d in enumerate(days)}
            for day in days:
                if day.dow not in ["Mon","Tue","Wed","Thu","Fri","Sat"]:
                    continue
                dslot = f"{day.date}-D"
                fslot = f"{day.date}-F"
                # if D/F slots don't exist for this day, skip
                v_d1 = x.get((dslot, doc1))
                v_d2 = x.get((dslot, doc2))
                v_f1 = x.get((fslot, doc1))
                v_f2 = x.get((fslot, doc2))
                if (v_d1 is None and v_d2 is None) or (v_f1 is None and v_f2 is None):
                    continue
                # Exception = (H is doc1/doc2) OR (J is doc1/doc2 oggi) OR (J is doc1/doc2 ieri)
                # Serve anche "J ieri": se doc2 ha fatto notte ieri, night_off next_day
                # gli vieta D oggi → il pattern non va imposto (altrimenti INFEASIBLE).
                conds = []
                hslot = f"{day.date}-H"
                jslot = f"{day.date}-J"
                for v in [x.get((hslot, doc1)), x.get((hslot, doc2)),
                          x.get((jslot, doc1)), x.get((jslot, doc2))]:
                    if v is not None:
                        conds.append(v)
                # J del giorno precedente
                _i_today = _day_idx_map.get(day.date)
                if _i_today is not None and _i_today > 0:
                    _prev_date = days[_i_today - 1].date
                    _prev_jslot = f"{_prev_date}-J"
                    for v in [x.get((_prev_jslot, doc1)), x.get((_prev_jslot, doc2))]:
                        if v is not None:
                            conds.append(v)
                if conds:
                    exc = model.NewBoolVar(f"exc_DF_{day.date}")
                    model.AddMaxEquality(exc, conds)  # exc = OR(conds)
                else:
                    exc = None
                if day.dow in ["Mon","Tue","Wed"]:
                    # Only enforce the 3-day pattern if BOTH required assignments are actually possible
                    # (i.e., corresponding variables exist after unavailability / domain filtering).
                    # Otherwise, enforcing only one side (e.g. F=doc2 while D=doc1 cannot happen)
                    # can make the whole month INFEASIBLE due to daily uniqueness constraints.
                    if v_d1 is None or v_f2 is None:
                        continue
                    # enforce D=doc1, F=doc2 when not exception
                    if v_d1 is not None:
                        if exc is None:
                            model.Add(v_d1 == 1)
                        else:
                            model.Add(v_d1 == 1).OnlyEnforceIf(exc.Not())
                    if v_f2 is not None:
                        if exc is None:
                            model.Add(v_f2 == 1)
                        else:
                            model.Add(v_f2 == 1).OnlyEnforceIf(exc.Not())
                else:
                    if v_d2 is None or v_f1 is None:
                        continue
                    # Thu/Fri/Sat: D=doc2, F=doc1 when not exception
                    if v_d2 is not None:
                        if exc is None:
                            model.Add(v_d2 == 1)
                        else:
                            model.Add(v_d2 == 1).OnlyEnforceIf(exc.Not())
                    if v_f1 is not None:
                        if exc is None:
                            model.Add(v_f1 == 1)
                        else:
                            model.Add(v_f1 == 1).OnlyEnforceIf(exc.Not())
    # Night off next day
    gc = cfg.get("global_constraints", {}) or {}
    night_off = (gc.get("night_off") or {})
    night_same = bool(night_off.get("same_day", True))
    night_next = bool(night_off.get("next_day", True))
    # Identify night slots
    night_slot_ids = [s.slot_id for s in slots if "J" in s.columns]
    # Night-off same day: if doctor works night(d) → no other slot(d) [except C which is exempt]
    if night_same:
        for drow in days:
            night_vars_day = []
            for sid in night_slot_ids:
                if sid.startswith(str(drow.date)):
                    for doc in doctors:
                        if (sid, doc) in x:
                            pass  # handled per-doc below
            for doc in doctors:
                night_vars = [x[(sid, doc)] for sid in night_slot_ids
                              if sid.startswith(str(drow.date)) and (sid, doc) in x]
                if not night_vars:
                    continue
                night_var = night_vars[0]
                # Every non-night, non-C slot same day → 0 if night=1
                for s2 in slots_by_day.get(drow.date, []):
                    if "J" in s2.columns:
                        continue  # skip the night slot itself
                    if any(c in {"C"} for c in (s2.columns or [])):
                        continue  # C is exempt
                    if (s2.slot_id, doc) in x:
                        model.Add(x[(s2.slot_id, doc)] == 0).OnlyEnforceIf(night_var)
    # Identify night slots
    night_slot_ids = [s.slot_id for s in slots if "J" in s.columns]
    if night_next:
        # for each day (except last), if doctor works night(d) then no slot(d+1)
        day_index = {d.date:i for i,d in enumerate(days)}
        for drow in days:
            i = day_index[drow.date]
            if i+1 >= len(days):
                continue
            next_day = days[i+1].date
            for doc in doctors:
                night_vars = []
                for sid in night_slot_ids:
                    if sid.startswith(str(drow.date)) and (sid, doc) in x:
                        night_vars.append(x[(sid, doc)])
                if not night_vars:
                    continue
                night_var = night_vars[0]  # only one night slot per day in our model
                # For every slot on next day, doc cannot be assigned if night_var=1
                for s2 in slots_by_day.get(next_day, []):
                    if (s2.slot_id, doc) in x:
                        model.Add(x[(s2.slot_id, doc)] == 0).OnlyEnforceIf(night_var)
    # Night spacing min
    min_gap = int(gc.get("night_spacing_days_min", 5))
    # Build night vars by (day,doc)
    night_var_by_day_doc = {}
    for s in slots:
        if "J" in s.columns:
            for doc in doctors:
                if (s.slot_id, doc) in x:
                    night_var_by_day_doc[(s.day.date, doc)] = x[(s.slot_id, doc)]
    # For each doc, prevent nights too close
    for doc in doctors:
        for i, drow in enumerate(days):
            for k in range(1, min_gap):
                j = i + k
                if j >= len(days):
                    break
                d1, d2 = drow.date, days[j].date
                v1 = night_var_by_day_doc.get((d1, doc))
                v2 = night_var_by_day_doc.get((d2, doc))
                if v1 is not None and v2 is not None:
                    model.Add(v1 + v2 <= 1)
    # Reperibilità C is reassigned after the CP-SAT solve by assign_reperibilita_C.
    # Do not let provisional C variables impose hard C↔J proximity constraints here:
    # they can make a real night J appear infeasible even though the final C layer
    # could simply choose a different reperibilità doctor.
    # K no consecutive days same doctor
    if "rules" in cfg and "K" in cfg["rules"] and cfg["rules"]["K"].get("no_consecutive_days_same_doctor", False):
        for i in range(len(days)-1):
            d1, d2 = days[i].date, days[i+1].date
            k1 = next((s for s in slots_by_day[d1] if "K" in s.columns), None)
            k2 = next((s for s in slots_by_day[d2] if "K" in s.columns), None)
            if not k1 or not k2:
                continue
            for doc in doctors:
                v1 = x.get((k1.slot_id, doc))
                v2 = x.get((k2.slot_id, doc))
                if v1 is not None and v2 is not None:
                    model.Add(v1 + v2 <= 1)
    
# D/F weekly behavior (Mon–Sat)
    # Goal (robust, NEVER infeasible):
    #   - Prefer Grimaldi + Calabrò on D/F when both are available
    #   - If only one of the pair is available: prefer D=that doctor; prefer F from H.pool_mon_fri
    #   - If none are available: prefer D and F from H.pool_mon_fri and prefer D=F (share)
    #
    # IMPORTANT: These are implemented as SOFT constraints with high penalties, so the model
    # can still remain FEASIBLE in "tight" days (lots of unavailability / other hard rules).
    if "rules" in cfg and "D_F" in cfg["rules"] and isinstance(cfg["rules"]["D_F"], dict):
        rDF = cfg["rules"]["D_F"]
        if rDF.get("pattern_3_3", False):
            doc1 = norm_name(rDF.get("pattern_doc1") or "Grimaldi")
            doc2 = norm_name(rDF.get("pattern_doc2") or "Calabrò")
            pair = {doc1, doc2}
            fallback_share_by_date: Dict[dt.date, Dict[str, object]] = {}
            fallback_share_by_doc: Dict[str, List[object]] = defaultdict(list)

            # Penalties (tunable in YAML)
            pen_pattern = int(rDF.get("pattern_penalty", 80) or 80)
            pen_outside_pair = int(rDF.get("outside_pair_penalty", 6000) or 6000)
            pen_missing_pair_doc = int(rDF.get("missing_pair_doc_penalty", 12000) or 12000)
            pen_outside_hpool = int(rDF.get("outside_hpool_penalty", 5000) or 5000)
            pen_fallback_balance = int(rDF.get("fallback_balance_penalty", 1_000_000) or 1_000_000)
            pen_fallback_switch = int(rDF.get("fallback_switch_penalty", 2_000_000) or 2_000_000)
            fallback_block_days = max(int(rDF.get("fallback_block_days", 3) or 3), 2)

            hpool = set(norm_name(d) for d in ((cfg.get("rules", {}).get("H", {}) or {}).get("pool_mon_fri") or []))

            for day in days:
                if day.dow not in ["Mon","Tue","Wed","Thu","Fri","Sat"]:
                    continue
                sD = next((s for s in slots_by_day[day.date] if s.columns == ["D"]), None)
                sF = next((s for s in slots_by_day[day.date] if s.columns == ["F"]), None)
                if not sD or not sF:
                    continue

                vD1 = x.get((sD.slot_id, doc1)); vF1 = x.get((sF.slot_id, doc1))
                vD2 = x.get((sD.slot_id, doc2)); vF2 = x.get((sF.slot_id, doc2))
                df_y_for_day = df_share_vars_by_date.get(day.date) or {}
                avail_pair = []
                if vD1 is not None and vF1 is not None:
                    avail_pair.append(doc1)
                if vD2 is not None and vF2 is not None:
                    avail_pair.append(doc2)

                # Helper: penalize choosing a doctor outside a set only if the set is actually feasible in-domain
                def penalize_outside(slot: Slot, allowed_set: set, penalty: int):
                    if not any((doc in allowed_set) and ((slot.slot_id, doc) in x) for doc in slot.allowed):
                        return
                    for doc in slot.allowed:
                        v = x.get((slot.slot_id, doc))
                        if v is not None and doc not in allowed_set:
                            extra_obj.append(penalty * v)

                # CASE: both pair doctors are available in-domain
                if len(avail_pair) == 2:
                    # Strongly prefer staying within the pair on BOTH slots
                    penalize_outside(sD, pair, pen_outside_pair)
                    penalize_outside(sF, pair, pen_outside_pair)

                    # Prefer that BOTH doctors appear across {D,F}
                    # (soft: if impossible due to other hard rules, solver may pay the penalty)
                    used1 = model.NewBoolVar(f"df_used_{day.date}_1")
                    used2 = model.NewBoolVar(f"df_used_{day.date}_2")
                    model.AddMaxEquality(used1, [v for v in [vD1, vF1] if v is not None])
                    model.AddMaxEquality(used2, [v for v in [vD2, vF2] if v is not None])
                    n1 = model.NewBoolVar(f"df_notused_{day.date}_1")
                    n2 = model.NewBoolVar(f"df_notused_{day.date}_2")
                    model.Add(n1 + used1 == 1)
                    model.Add(n2 + used2 == 1)
                    extra_obj.append(n1 * pen_missing_pair_doc)
                    extra_obj.append(n2 * pen_missing_pair_doc)

                    # Weekly pattern preference (only among the pair)
                    prefD, prefF = (doc1, doc2) if day.dow in ["Mon","Tue","Wed"] else (doc2, doc1)
                    for doc in pair:
                        v = x.get((sD.slot_id, doc))
                        if v is not None and doc != prefD:
                            extra_obj.append(pen_pattern * v)
                    for doc in pair:
                        v = x.get((sF.slot_id, doc))
                        if v is not None and doc != prefF:
                            extra_obj.append(pen_pattern * v)

                # CASE: only one of the pair is available in-domain
                elif len(avail_pair) == 1:
                    only_doc = avail_pair[0]
                    vD_only = x.get((sD.slot_id, only_doc))
                    vF_only = x.get((sF.slot_id, only_doc))

                    # HARD: if exactly one primary doctor is available in both D and F,
                    # that doctor covers the pair. `pair_avail` already excludes
                    # unavailability, forced J same-day and Saturday Calabro cases.
                    if vD_only is not None and vF_only is not None:
                        model.Add(vD_only == 1)
                        model.Add(vF_only == 1)

                    # NON penalizzare altri medici in F (l'unico del pair li occupa entrambi)

                # CASE: none of the pair is available in-domain
                else:
                    # HARD: D e F devono essere assegnati allo stesso medico dell'H-pool.
                    # Usiamo il vincolo df_share già costruito sopra (df_y_by_doc) per forzare la share.
                    # Penalizziamo solo dottori fuori dall'H-pool per orientare la scelta.
                    penalize_outside(sD, hpool, pen_outside_hpool)
                    penalize_outside(sF, hpool, pen_outside_hpool)
                    # Forza D=F tramite il meccanismo df_share già presente
                    if df_y_for_day:
                        share_any = model.NewBoolVar(f"df_share_none_{day.date.isoformat()}")
                        model.AddMaxEquality(share_any, list(df_y_for_day.values()))
                        model.Add(share_any == 1)  # HARD: deve esserci un medico che copre sia D che F
                        fallback_share_by_date[day.date] = dict(df_y_for_day)
                        for _doc, _share_var in df_y_for_day.items():
                            fallback_share_by_doc[_doc].append(_share_var)
                    else:
                        pre_solve_warnings.append(
                            f"D/F il {day.date}: nessun medico disponibile per coprire entrambe le colonne (D e F lasciate scoperte)"
                        )
            # When D/F falls back because both primary doctors are unavailable,
            # balance fallback doctors and prefer stable consecutive blocks.
            if fallback_share_by_doc:
                loads = []
                max_prior = max((_prior_rule_balance_count(d, "D_F_fallback") for d in fallback_share_by_doc), default=0)
                for _doc, _vars in sorted(fallback_share_by_doc.items()):
                    if not _vars:
                        continue
                    prior = _prior_rule_balance_count(_doc, "D_F_fallback")
                    load = model.NewIntVar(0, len(_vars) + prior, f"df_fb_load_{hash(_doc)%10**6}")
                    model.Add(load == sum(_vars) + prior)
                    loads.append(load)
                if loads:
                    max_fb = model.NewIntVar(0, len(fallback_share_by_date) + max_prior, "df_fb_max_load")
                    min_fb = model.NewIntVar(0, len(fallback_share_by_date) + max_prior, "df_fb_min_load")
                    spread_fb = model.NewIntVar(0, len(fallback_share_by_date) + max_prior, "df_fb_spread")
                    model.AddMaxEquality(max_fb, loads)
                    model.AddMinEquality(min_fb, loads)
                    model.Add(spread_fb == max_fb - min_fb)
                    extra_obj.append(pen_fallback_balance * (max_fb + spread_fb))

                fb_dates = sorted(fallback_share_by_date)
                runs: List[List[dt.date]] = []
                for _date in fb_dates:
                    if not runs or (runs[-1][-1] + dt.timedelta(days=1)) != _date:
                        runs.append([_date])
                    else:
                        runs[-1].append(_date)
                for run in runs:
                    for i in range(0, len(run), fallback_block_days):
                        chunk = run[i:i + fallback_block_days]
                        if len(chunk) < 2:
                            continue
                        for d1, d2 in zip(chunk, chunk[1:]):
                            m1 = fallback_share_by_date.get(d1) or {}
                            m2 = fallback_share_by_date.get(d2) or {}
                            common_docs = sorted(set(m1) & set(m2))
                            same_terms = []
                            for _doc in common_docs:
                                same = model.NewBoolVar(f"df_fb_same_{d1}_{d2}_{hash(_doc)%10**6}")
                                model.Add(same <= m1[_doc])
                                model.Add(same <= m2[_doc])
                                model.Add(same >= m1[_doc] + m2[_doc] - 1)
                                same_terms.append(same)
                            if same_terms:
                                switch = model.NewBoolVar(f"df_fb_switch_{d1}_{d2}")
                                model.Add(sum(same_terms) + switch == 1)
                                extra_obj.append(pen_fallback_switch * switch)
# E/G weekly blocks (Mon-Sat) if block_days=6
    if "rules" in cfg and "E_G" in cfg["rules"]:
        block_days = int(cfg["rules"]["E_G"].get("block_days", 0) or 0)
        # Find all EG slots by date
        eg_by_date = {}
        for s in slots:
            if s.columns == ["E","G"]:
                eg_by_date[s.day.date] = s
        if block_days == 6:
            # For each Monday, enforce same doctor Mon..Sat (if all present)
            for day in days:
                if day.dow != "Mon":
                    continue
                seq = [day.date + dt.timedelta(days=k) for k in range(0,6)]
                if not all(d in eg_by_date for d in seq):
                    continue
                for doc in doctors:
                    v0 = x.get((eg_by_date[seq[0]].slot_id, doc))
                    if v0 is None:
                        continue
                    for d in seq[1:]:
                        v = x.get((eg_by_date[d].slot_id, doc))
                        if v is not None:
                            model.Add(v == v0)
        elif block_days == 3:
            # Split week into 2 blocks: Mon-Wed and Thu-Sat (if all present).
            for day in days:
                if day.dow != "Mon":
                    continue
                seq1 = [day.date + dt.timedelta(days=k) for k in range(0,3)]
                seq2 = [day.date + dt.timedelta(days=k) for k in range(3,6)]
                if not all(d in eg_by_date for d in (seq1 + seq2)):
                    continue
                for doc in doctors:
                    v0 = x.get((eg_by_date[seq1[0]].slot_id, doc))
                    if v0 is not None:
                        for d in seq1[1:]:
                            v = x.get((eg_by_date[d].slot_id, doc))
                            if v is not None:
                                model.Add(v == v0)
                for doc in doctors:
                    v3 = x.get((eg_by_date[seq2[0]].slot_id, doc))
                    if v3 is not None:
                        for d in seq2[1:]:
                            v = x.get((eg_by_date[d].slot_id, doc))
                            if v is not None:
                                model.Add(v == v3)
                # Soft preference: use two different doctors between the two 3-day blocks.
                # Penalize if the SAME doctor is chosen on Mon (block1 start) and Thu (block2 start).
                pen_split = int((cfg["rules"]["E_G"].get("block_split_penalty", 10) or 10))
                mon_slot = eg_by_date[seq1[0]]
                thu_slot = eg_by_date[seq2[0]]
                same_terms = []
                for doc in doctors:
                    vmon = x.get((mon_slot.slot_id, doc))
                    vthu = x.get((thu_slot.slot_id, doc))
                    if vmon is None or vthu is None:
                        continue
                    b = model.NewBoolVar(f"eg_same_{hash(doc)%10**6}_{day.date}")
                    model.AddBoolAnd([vmon, vthu]).OnlyEnforceIf(b)
                    model.AddBoolOr([vmon.Not(), vthu.Not()]).OnlyEnforceIf(b.Not())
                    same_terms.append(b)
                if same_terms:
                    extra_obj.append(pen_split * sum(same_terms))

        # E/G: bilanciamento hard — ogni medico può fare al massimo ceil(slots/pool_attivo)+1 blocchi
        eg_slots_all = list(eg_by_date.values())
        if eg_slots_all:
            rEG = cfg["rules"]["E_G"]
            eg_pool = [norm_name(d) for d in (rEG.get("allowed") or []) if norm_name(d) in doctors]
            if eg_pool:
                n_eg = len(eg_slots_all)
                # Conta solo i medici con almeno una variabile disponibile (tiene conto delle indisponibilità)
                eg_active_docs = [
                    d for d in eg_pool
                    if any(x.get((s.slot_id, d)) is not None for s in eg_slots_all)
                ]
                n_pool_active = max(len(eg_active_docs), 1)
                eg_max_hard = math.ceil(n_eg / n_pool_active) + 1  # basato su pool attivo
                eg_cnt_vars = []
                for doc in eg_pool:
                    vars_ = [x.get((s.slot_id, doc)) for s in eg_slots_all if x.get((s.slot_id, doc)) is not None]
                    if vars_:
                        eg_cnt = model.NewIntVar(0, n_eg, f"eg_cnt_{hash(doc)%10**6}")
                        model.Add(eg_cnt == sum(vars_))
                        # Soft: penalizza eccesso (no hard per evitare infeasibility)
                        _eg_over = model.NewIntVar(0, n_eg, f"eg_over_{hash(doc)%10**6}")
                        model.Add(_eg_over >= eg_cnt - eg_max_hard)
                        model.Add(_eg_over >= 0)
                        extra_obj.append(300_000 * _eg_over)
                        eg_cnt_vars.append(eg_cnt)
                if eg_cnt_vars:
                    eg_max_v = model.NewIntVar(0, n_eg, "eg_max_load")
                    model.AddMaxEquality(eg_max_v, eg_cnt_vars)
                    extra_obj.append(300 * eg_max_v)  # forte penalità per minimizzare il massimo
    # Monthly quotas — J
    # Zito, Dattilo, Calabrò: quota ESATTA (hard upper + penalità 25M deficit).
    # Licordari, Colarusso e gli altri: nessuna quota fissa qui — bilanciamento
    # soft [min_per, max_per] gestito più sotto (target 2, fino a 3 se necessario).
    J_QUOTA_DEV_PENALTY = 8_000_000   # default
    J_QUOTA_STRICT_PENALTY = 25_000_000  # Zito, Dattilo, Calabrò — quasi-obbligatorio
    _j_strict_docs = {norm_name("Zito"), norm_name("Dattilo"), norm_name("Calabrò")}
    if "rules" in cfg and "J" in cfg["rules"]:
        mq = cfg["rules"]["J"].get("monthly_quotas") or {}
        for doc_raw, q in mq.items():
            doc = norm_name(doc_raw)
            if doc not in doctors:
                continue
            vars_ = [night_var_by_day_doc.get((d.date, doc)) for d in days]
            vars_ = [v for v in vars_ if v is not None]
            if not vars_:
                continue
            q_int = int(q)
            n_avail = len(vars_)
            q_target = min(_monthly_target_for_period(q_int, doc, "J"), n_avail)
            q_cap = min(_monthly_cap_for_period(q_int, doc, "J"), n_avail)
            q_cap = max(q_cap, q_target)
            if n_avail < q_int and not period_is_partial_month:
                pre_solve_warnings.append(
                    f"J quota {doc}: richieste {q_int} notti ma solo {n_avail} disponibili. "
                    f"Quota adattata a {q_target}."
                )
            _strict = doc in _j_strict_docs
            _pen = J_QUOTA_STRICT_PENALTY if _strict else J_QUOTA_DEV_PENALTY
            # Hard upper per i medici a quota stretta (Zito, Dattilo, Calabrò):
            # non devono mai superare la loro quota.
            if _strict:
                model.Add(sum(vars_) <= q_cap)
                if q_cap > q_target:
                    _jsup = model.NewIntVar(0, n_avail, f"jsup_strict_{hash(doc)%10**6}")
                    model.Add(_jsup >= sum(vars_) - q_target)
                    model.Add(_jsup >= 0)
                    extra_obj.append(_pen * _jsup)
            else:
                # Soft surplus per gli altri (hard upper già gestito da pool_quota_overrides)
                _jsup = model.NewIntVar(0, n_avail, f"jsup_{hash(doc)%10**6}")
                model.Add(_jsup >= sum(vars_) - q_target)
                model.Add(_jsup >= 0)
                extra_obj.append(_pen * _jsup)
            # Soft deficit: penalità per ogni notte mancante rispetto alla quota
            _jsum = model.NewIntVar(0, n_avail, f"jsum_{hash(doc)%10**6}")
            model.Add(_jsum == sum(vars_))
            _jdef = model.NewIntVar(0, n_avail, f"jdef_{hash(doc)%10**6}")
            model.Add(_jdef >= q_target - _jsum)
            model.Add(_jdef >= 0)
            extra_obj.append(_pen * _jdef)
    # Pool quota overrides max/min — da pool_config (tutti i tipi e colonne)
    qov = cfg.get("pool_quota_overrides") or {}
    for (doc_n, col), spec in qov.items():
        col_slots = [s for s in slots if col in (s.columns or [])]
        vars_ = [x.get((s.slot_id, doc_n)) for s in col_slots]
        vars_ = [v for v in vars_ if v is not None]
        if not vars_:
            continue
        sv = sum(vars_)
        qt = spec.get("type", "max")
        val_monthly = int(spec.get("value", 0))
        n_avail = len(vars_)
        if qt == "max":
            model.Add(sv <= min(_monthly_cap_for_period(val_monthly, doc_n, col), n_avail))
        elif qt == "min":
            effective_min = min(_monthly_target_for_period(val_monthly, doc_n, col), n_avail)
            if effective_min > 0:
                model.Add(sv >= effective_min)
            if effective_min < val_monthly and not period_is_partial_month:
                pre_solve_warnings.append(
                    f"Quota min {doc_n}/{col}: richiesto min {val_monthly} ma solo {n_avail} slot disponibili."
                )
        elif qt == "fixed":
            effective_target = min(_monthly_target_for_period(val_monthly, doc_n, col), n_avail)
            effective_cap = min(_monthly_cap_for_period(val_monthly, doc_n, col), n_avail)
            effective_cap = max(effective_cap, effective_target)
            # Hard upper bound + soft deviation from the period target.
            model.Add(sv <= effective_cap)
            _fdef = model.NewIntVar(0, n_avail, f"fdef_{hash((doc_n,col))%10**6}")
            model.Add(_fdef >= effective_target - sv)
            model.Add(_fdef >= 0)
            extra_obj.append(J_QUOTA_DEV_PENALTY * _fdef)
            if effective_cap > effective_target:
                _fsup = model.NewIntVar(0, n_avail, f"fsup_{hash((doc_n,col))%10**6}")
                model.Add(_fsup >= sv - effective_target)
                model.Add(_fsup >= 0)
                extra_obj.append(J_QUOTA_DEV_PENALTY * _fsup)
            if effective_target < val_monthly and not period_is_partial_month:
                pre_solve_warnings.append(
                    f"Quota fixed {doc_n}/{col}: richiesto {val_monthly} ma solo {n_avail} slot disponibili, ridotto a {effective_target}."
                )

    # Target mensili globali da GUI: obiettivo soft per tutti i medici del pool
    # della colonna. Gli override per singolo medico prevalgono e vengono saltati.
    pool_monthly_targets = cfg.get("pool_monthly_targets") or {}
    if pool_monthly_targets:
        _rule_for_col = {
            "D": "D_F", "F": "D_F", "E": "E_G", "G": "E_G",
            "H": "H", "I": "I", "J": "J", "K": "K", "L": "L",
            "Q": "Q", "R": "R", "S": "S", "T": "T", "U": "U",
            "V": "V", "W": "W", "Y": "Y", "Z": "Z", "AB": "AB",
        }
        _j_mq_docs = {
            norm_name(d)
            for d in ((cfg.get("rules") or {}).get("J", {}).get("monthly_quotas") or {}).keys()
        }
        for col_raw, target_raw in pool_monthly_targets.items():
            col = str(col_raw).strip().upper()
            if col == "C":
                continue
            try:
                target_monthly = int(target_raw)
            except Exception:
                continue
            if target_monthly < 0:
                continue
            col_slots = [s for s in slots if col in (s.columns or [])]
            if not col_slots:
                continue
            rcol = (cfg.get("rules") or {}).get(_rule_for_col.get(col, col), {}) or {}
            target_penalty = max(int(rcol.get("balance_weight") or 200), 1) * 250
            for doc in doctors:
                if (doc, col) in qov:
                    continue
                if col == "J" and doc in _j_mq_docs:
                    continue
                vars_ = [x.get((s.slot_id, doc)) for s in col_slots]
                vars_ = [v for v in vars_ if v is not None]
                if not vars_:
                    continue
                n_avail = len(vars_)
                target = _monthly_target_for_period(target_monthly, doc, col)
                effective_target = min(target, n_avail)
                cnt = model.NewIntVar(0, n_avail, f"mt_cnt_{hash((doc,col))%10**6}")
                model.Add(cnt == sum(vars_))
                under = model.NewIntVar(0, effective_target, f"mt_under_{hash((doc,col))%10**6}")
                over = model.NewIntVar(0, n_avail, f"mt_over_{hash((doc,col))%10**6}")
                model.Add(under >= effective_target - cnt)
                model.Add(under >= 0)
                model.Add(over >= cnt - effective_target)
                model.Add(over >= 0)
                extra_obj.append(target_penalty * (under + over))
    # Monthly quotas (hard) — Festivi DE+HI
    if "rules" in cfg and "Festivi" in cfg["rules"]:
        rFest = cfg["rules"]["Festivi"]
        fest_quotas = {norm_name(k): int(v) for k, v in (rFest.get("quotas") or {}).items()}
        if fest_quotas:
            festivo_slots = [s for s in slots if s.rule_tag in ("Festivo_DE", "Festivo_HI")]
            for doc, q in fest_quotas.items():
                if doc not in doctors:
                    continue
                vars_ = [x.get((s.slot_id, doc)) for s in festivo_slots]
                vars_ = [v for v in vars_ if v is not None]
                if vars_:
                    q_target = min(_monthly_target_for_period(q, doc, "Festivi"), len(vars_))
                    q_cap = min(_monthly_cap_for_period(q, doc, "Festivi"), len(vars_))
                    q_cap = max(q_cap, q_target)
                    # Hard upper bound + soft deviation from the period target.
                    model.Add(sum(vars_) <= q_cap)
                    _fqdef = model.NewIntVar(0, len(vars_), f"fqdef_{hash(doc)%10**6}")
                    model.Add(_fqdef >= q_target - sum(vars_))
                    model.Add(_fqdef >= 0)
                    extra_obj.append(J_QUOTA_DEV_PENALTY * _fqdef)
                    if q_cap > q_target:
                        _fqsup = model.NewIntVar(0, len(vars_), f"fqsup_{hash(doc)%10**6}")
                        model.Add(_fqsup >= sum(vars_) - q_target)
                        model.Add(_fqsup >= 0)
                        extra_obj.append(J_QUOTA_DEV_PENALTY * _fqsup)
                    if q_target < q and not period_is_partial_month:
                        pre_solve_warnings.append(
                            f"Festivi quota {doc}: richiesti {q} ma solo {len(vars_)} slot disponibili."
                        )
    # Soft balance festivi — bilancia D+E e H+I insieme.
    # Le domeniche/festivi diurni vengono percepiti come un unico carico: non
    # basta bilanciare D/E separatamente da H/I.
    if "rules" in cfg and "Festivi" in cfg["rules"]:
        try:
            rFest = cfg["rules"]["Festivi"]
            festivo_slots_all = [s for s in slots if s.rule_tag in ("Festivo_DE", "Festivo_HI")]
            balance_pool = sorted({
                doc
                for s in festivo_slots_all
                for doc in (s.allowed or [])
                if doc in doctors and doc != "Recupero"
            })
            fest_bal_w = int(
                rFest.get("sunday_de_hi_balance_penalty")
                or rFest.get("festivi_diurni_spread_penalty")
                or 25_000_000
            )
            if balance_pool and festivo_slots_all:
                fest_quota_penalty = int(rFest.get("festivi_diurni_quota_penalty") or 250_000_000)
                fest_concentration_penalty = int(rFest.get("festivi_diurni_concentration_penalty") or 90_000_000)
                fest_fixed_docs = {norm_name(k) for k in (rFest.get("quotas") or {}).keys()}
                free_fest_docs = [d for d in balance_pool if d not in fest_fixed_docs]
                if fest_quota_penalty > 0 and free_fest_docs:
                    fixed_total = sum(
                        _monthly_target_for_period(int(v), norm_name(doc), "Festivi")
                        for doc, v in (rFest.get("quotas") or {}).items()
                        if norm_name(doc) in doctors
                    )
                    free_total = max(0, len(festivo_slots_all) - fixed_total)
                    n_free = len(free_fest_docs)
                    min_per = free_total // n_free
                    remainder = free_total - min_per * n_free
                    max_per = min_per + (1 if remainder > 0 else 0)
                    for d in free_fest_docs:
                        vars_d = [
                            x[(s.slot_id, d)]
                            for s in festivo_slots_all
                            if (s.slot_id, d) in x
                        ]
                        if not vars_d:
                            continue
                        prior_fest = _prior_count(d, "Festivi")
                        total_upper = len(vars_d) + prior_fest
                        if fest_concentration_penalty > 0:
                            fest_cnt = model.NewIntVar(0, total_upper, f"fest_cnt_{hash(d)%10**6}")
                            model.Add(fest_cnt == sum(vars_d) + prior_fest)
                            for threshold in range(2, total_upper + 1):
                                over_threshold = model.NewIntVar(
                                    0,
                                    total_upper,
                                    f"fest_conc_{threshold}_{hash(d)%10**6}",
                                )
                                model.Add(over_threshold >= fest_cnt - (threshold - 1))
                                model.Add(over_threshold >= 0)
                                extra_obj.append(fest_concentration_penalty * threshold * over_threshold)
                        cur_min = max(0, min_per - prior_fest)
                        cur_max = max(0, max_per - prior_fest)
                        cur_sum = sum(vars_d)
                        if cur_min > 0:
                            under = model.NewIntVar(0, len(vars_d), f"fest_under_{hash(d)%10**6}")
                            model.Add(under >= cur_min - cur_sum)
                            model.Add(under >= 0)
                            extra_obj.append(fest_quota_penalty * under)
                        over = model.NewIntVar(0, len(vars_d), f"fest_over_{hash(d)%10**6}")
                        model.Add(over >= cur_sum - cur_max)
                        model.Add(over >= 0)
                        extra_obj.append(fest_quota_penalty * over)

                max_prior_fest = max((_prior_count(d, "Festivi") for d in balance_pool), default=0)
                fest_loads = []
                for d in balance_pool:
                    vars_d = [x[(s.slot_id, d)] for s in festivo_slots_all
                              if (s.slot_id, d) in x]
                    if not vars_d:
                        continue
                    prior_fest = _prior_count(d, "Festivi")
                    load_d = model.NewIntVar(0, len(festivo_slots_all) + prior_fest,
                                            f"fest_load_{hash(d) % 10**6}")
                    model.Add(load_d == sum(vars_d) + prior_fest)
                    fest_loads.append(load_d)
                if len(fest_loads) > 1:
                    max_fest = model.NewIntVar(0, len(festivo_slots_all) + max_prior_fest, "max_fest_load")
                    min_fest = model.NewIntVar(0, len(festivo_slots_all) + max_prior_fest, "min_fest_load")
                    diff_fest = model.NewIntVar(0, len(festivo_slots_all) + max_prior_fest, "diff_fest_load")
                    model.AddMaxEquality(max_fest, fest_loads)
                    model.AddMinEquality(min_fest, fest_loads)
                    model.Add(diff_fest == max_fest - min_fest)
                    extra_obj.append(fest_bal_w * diff_fest)
        except Exception:
            pass
    # Soft: alcuni medici devono preferibilmente avere almeno N notti weekend (sab/dom)
    if "rules" in cfg and "J" in cfg["rules"]:
        rJ_wn = cfg["rules"]["J"]
        wn_min_soft = rJ_wn.get("weekend_night_min_soft") or {}
        wn_pen = int(rJ_wn.get("weekend_night_min_soft_penalty") or 3000)
        for doc_raw, min_we in wn_min_soft.items():
            doc = norm_name(doc_raw)
            if doc not in doctors:
                continue
            we_vars = [night_var_by_day_doc.get((d.date, doc))
                       for d in days if d.dow in ("Sat", "Sun")]
            we_vars = [v for v in we_vars if v is not None]
            if not we_vars:
                continue  # Zito indisponibile tutti i weekend: vincolo ignorato
            min_we = _monthly_target_for_period(int(min_we), doc, "J")
            if min_we <= 0:
                continue
            no_we = model.NewBoolVar(f"no_we_night_{hash(doc)%10**6}")
            we_sum = model.NewIntVar(0, len(we_vars), f"we_sum_{hash(doc)%10**6}")
            model.Add(we_sum == sum(we_vars))
            model.Add(we_sum < min_we).OnlyEnforceIf(no_we)
            model.Add(we_sum >= min_we).OnlyEnforceIf(no_we.Not())
            extra_obj.append(wn_pen * no_we)
        # Hard max weekend nights per dottore (es. Zito: max 1)
        wn_max_hard = rJ_wn.get("weekend_night_max_hard") or {}
        for doc_raw, max_we in wn_max_hard.items():
            doc = norm_name(doc_raw)
            if doc not in doctors:
                continue
            we_vars = [night_var_by_day_doc.get((d.date, doc))
                       for d in days if d.dow in ("Sat", "Sun")]
            we_vars = [v for v in we_vars if v is not None]
            if we_vars:
                model.Add(sum(we_vars) <= _monthly_cap_for_period(int(max_we), doc, "J"))
    # Night distribution (HARD min/max per dottore + soft balance weekend)
    # Logica: total_nights = giorni del mese - giovedì (thursday_blank).
    # pool_available_nights esclude slot J pre-assegnate a medici fuori pool (es. festivi fissi).
    # Quota fissa (monthly_quotas YAML) sottratta → free_total diviso equamente tra free_docs.
    # Regola generale: ogni free doctor fa MIN floor(free_total/n), MAX floor+1 notti.
    if "rules" in cfg and "J" in cfg["rules"]:
        rJ = cfg["rules"]["J"]
        night_pool = set(norm_name(d) for d in (rJ.get("pool_other") or []))
        night_pool |= set(norm_name(d) for d in (rJ.get("monthly_quotas") or {}).keys())
        night_pool = {d for d in night_pool if d in doctors and d != "Recupero"}
        mq_fixed = {norm_name(k): int(v) for k,v in (rJ.get("monthly_quotas") or {}).items()
                    if norm_name(k) in doctors}
        total_nights = sum(1 for s in slots if s.columns == ["J"])
        # Notti disponibili per i pool doctors (esclude slot pre-assegnate a medici fuori pool)
        pool_available_nights = sum(
            1 for s in slots
            if s.columns == ["J"] and any(d in night_pool for d in s.allowed)
        )

        if night_pool and total_nights > 0:
            # Medici con quota fissa: già vincolati con == sopra.
            # Medici senza quota fissa: imponiamo min=2, max=3 hard.
            free_docs = [d for d in sorted(night_pool) if d not in mq_fixed]
            fixed_total = sum(_monthly_target_for_period(v, doc, "J") for doc, v in mq_fixed.items())
            free_total = max(0, pool_available_nights - fixed_total)

            # Calcola min/max bilanciati per i medici liberi
            if free_docs:
                n_free = len(free_docs)
                prior_free_total = sum(_prior_count(doc, "J") for doc in free_docs)
                month_free_total = free_total + prior_free_total
                # month_free_total / n_free → es. 21/9 = 2.33 → min=2, max=3.
                # Nei periodi parziali successivi la prima settimana salvata resta
                # nel totale mensile tramite prior_free_total.
                min_per = month_free_total // n_free  # minimo garantito sul mese logico
                remainder = month_free_total - min_per * n_free
                # max_per = min_per se il resto è 0, altrimenti min_per+1
                max_per = min_per + (1 if remainder > 0 else 0)
                max_per = max(max_per, 0)  # sicurezza: mai negativo
                free_quota_penalty = int(rJ.get("free_night_quota_penalty", 30_000_000) or 0)
                free_spread_penalty = int(rJ.get("free_night_spread_penalty", 20_000_000) or 0)

                for doc in free_docs:
                    vars_ = [night_var_by_day_doc.get((d.date, doc)) for d in days
                             if night_var_by_day_doc.get((d.date, doc)) is not None]
                    if vars_:
                        prior_j = _prior_count(doc, "J")
                        cur_min = max(0, min_per - prior_j)
                        cur_max = max(0, max_per - prior_j)
                        cur_sum = sum(vars_)
                        if free_quota_penalty > 0:
                            if cur_min > 0:
                                under_free = model.NewIntVar(0, len(vars_), f"j_free_under_{hash(doc)%10**6}")
                                model.Add(under_free >= cur_min - cur_sum)
                                model.Add(under_free >= 0)
                                extra_obj.append(free_quota_penalty * under_free)
                            over_free = model.NewIntVar(0, len(vars_), f"j_free_over_{hash(doc)%10**6}")
                            model.Add(over_free >= cur_sum - cur_max)
                            model.Add(over_free >= 0)
                            extra_obj.append(free_quota_penalty * over_free)

                # Soft balance: minimizza la differenza max-min tra i medici liberi.
                # Deve avere un peso reale: con pesi bassi il solver può accettare
                # distribuzioni 3/3/1/1 pur di ottimizzare dettagli secondari.
                if free_spread_penalty > 0 and len(free_docs) > 1:
                    cnt_vars = []
                    for doc in free_docs:
                        vars_ = [night_var_by_day_doc.get((d.date, doc)) for d in days
                                 if night_var_by_day_doc.get((d.date, doc)) is not None]
                        if vars_:
                            cnt = model.NewIntVar(0, month_free_total, f"nightcnt_{hash(doc)%10**6}")
                            model.Add(cnt == sum(vars_) + _prior_count(doc, "J"))
                            cnt_vars.append(cnt)
                    if cnt_vars:
                        max_cnt = model.NewIntVar(0, month_free_total, "night_max_free")
                        min_cnt = model.NewIntVar(0, month_free_total, "night_min_free")
                        model.AddMaxEquality(max_cnt, cnt_vars)
                        model.AddMinEquality(min_cnt, cnt_vars)
                        diff_cnt = model.NewIntVar(0, month_free_total, "night_diff_free")
                        model.Add(diff_cnt == max_cnt - min_cnt)
                        extra_obj.append(free_spread_penalty * diff_cnt)

        # Weekday J balance: separate from weekend nights. The total J balance
        # alone can hide an uneven split between Mon-Fri nights and Sat/Sun.
        weekday_spread_penalty = int(rJ.get("weekday_night_spread_penalty", 25_000_000) or 0)
        if weekday_spread_penalty > 0:
            weekday_cnt_vars = []
            max_prior_weekday = max((_prior_j_weekday_count(doc) for doc in night_pool), default=0)
            weekday_upper = sum(1 for day in days if not _day_is_j_festive(day)) + max_prior_weekday
            for doc in sorted(night_pool):
                weekday_vars = []
                for day in days:
                    if not _day_is_j_festive(day):
                        v = night_var_by_day_doc.get((day.date, doc))
                        if v is not None:
                            weekday_vars.append(v)
                if not weekday_vars:
                    continue
                prior_weekday_doc = _prior_j_weekday_count(doc)
                weekday_cnt = model.NewIntVar(
                    0,
                    len(weekday_vars) + prior_weekday_doc,
                    f"weekday_j_{hash(doc)%10**6}",
                )
                model.Add(weekday_cnt == sum(weekday_vars) + prior_weekday_doc)
                weekday_cnt_vars.append(weekday_cnt)
            if len(weekday_cnt_vars) > 1:
                weekday_max = model.NewIntVar(0, max(weekday_upper, 1), "weekday_j_max")
                weekday_min = model.NewIntVar(0, max(weekday_upper, 1), "weekday_j_min")
                weekday_diff = model.NewIntVar(0, max(weekday_upper, 1), "weekday_j_diff")
                model.AddMaxEquality(weekday_max, weekday_cnt_vars)
                model.AddMinEquality(weekday_min, weekday_cnt_vars)
                model.Add(weekday_diff == weekday_max - weekday_min)
                extra_obj.append(weekday_spread_penalty * weekday_diff)

        # Weekend nights: the generic fair-share cap must be SOFT. In vacation
        # periods it is common that the only available Sunday-night candidates
        # already have one weekend night; leaving J blank is worse than breaking
        # the generic cap. Explicit per-doctor caps above (e.g. Zito) remain hard.
        _j_wex = {norm_name(d) for d in (rJ.get("weekend_excluded_doctors") or ["Calabrò"])}
        weekend_docs = night_pool - _j_wex
        import math as _math
        total_we_nights = sum(
            1 for day in days if _day_is_j_festive(day)
            if any(night_var_by_day_doc.get((day.date, doc)) is not None for doc in weekend_docs)
        )
        prior_we_nights = sum(_prior_j_weekend_count(doc) for doc in weekend_docs)
        month_we_nights = total_we_nights + prior_we_nights
        n_we_docs = len(weekend_docs)
        we_hard_cap = _math.ceil(month_we_nights / n_we_docs) if n_we_docs > 0 else 1
        we_soft_target = (month_we_nights // n_we_docs) if n_we_docs > 0 else 1
        we_cap_penalty = int(rJ.get("weekend_night_cap_penalty", 30_000_000) or 0)
        we_spread_penalty = int(rJ.get("weekend_night_spread_penalty", 30_000_000) or 0)
        we_cnt_vars = []
        for doc in sorted(weekend_docs):
            we_vars = []
            for day in days:
                if _day_is_j_festive(day):
                    v = night_var_by_day_doc.get((day.date, doc))
                    if v is not None:
                        we_vars.append(v)
            if we_vars:
                prior_we_doc = _prior_j_weekend_count(doc)
                we_cnt = model.NewIntVar(0, len(we_vars) + prior_we_doc, f"we_night_{hash(doc)%10**6}")
                model.Add(we_cnt == sum(we_vars) + prior_we_doc)
                if we_cap_penalty > 0:
                    over_cap = model.NewIntVar(0, len(we_vars) + prior_we_doc, f"we_over_cap_{hash(doc)%10**6}")
                    model.Add(over_cap >= we_cnt - we_hard_cap)
                    model.Add(over_cap >= 0)
                    extra_obj.append(we_cap_penalty * over_cap)
                we_cnt_vars.append(we_cnt)
                if we_hard_cap > we_soft_target:
                    over_tgt = model.NewIntVar(0, len(we_vars) + prior_we_doc, f"we_over_tgt_{hash(doc)%10**6}")
                    model.Add(over_tgt >= we_cnt - we_soft_target)
                    model.Add(over_tgt >= 0)
                    extra_obj.append(max(we_spread_penalty // 4, 1) * over_tgt)
        if we_cnt_vars:
            we_upper = max(month_we_nights, 1)
            we_max = model.NewIntVar(0, we_upper, "we_night_max")
            model.AddMaxEquality(we_max, we_cnt_vars)
            if len(we_cnt_vars) > 1:
                we_min = model.NewIntVar(0, we_upper, "we_night_min")
                model.AddMinEquality(we_min, we_cnt_vars)
                we_diff = model.NewIntVar(0, we_upper, "we_night_diff")
                model.Add(we_diff == we_max - we_min)
                if we_spread_penalty > 0:
                    extra_obj.append(we_spread_penalty * we_diff)

        # Per-doctor J type split: for the same total number of J, prefer a
        # balanced split between feriali and festive/weekend nights. This makes
        # 2+2 better than 3+1 when both are feasible.
        type_split_penalty = int(rJ.get("j_type_split_penalty", 35_000_000) or 0)
        if type_split_penalty > 0:
            for doc in sorted(weekend_docs):
                weekday_vars = []
                festive_vars = []
                for day in days:
                    v = night_var_by_day_doc.get((day.date, doc))
                    if v is None:
                        continue
                    if _day_is_j_festive(day):
                        festive_vars.append(v)
                    else:
                        weekday_vars.append(v)
                if not weekday_vars or not festive_vars:
                    continue
                prior_weekday_doc = _prior_j_weekday_count(doc)
                prior_festive_doc = _prior_j_weekend_count(doc)
                weekday_cnt = model.NewIntVar(
                    0,
                    len(weekday_vars) + prior_weekday_doc,
                    f"jtype_weekday_{hash(doc)%10**6}",
                )
                festive_cnt = model.NewIntVar(
                    0,
                    len(festive_vars) + prior_festive_doc,
                    f"jtype_festive_{hash(doc)%10**6}",
                )
                split_upper = len(weekday_vars) + len(festive_vars) + prior_weekday_doc + prior_festive_doc
                split_diff = model.NewIntVar(0, max(split_upper, 1), f"jtype_diff_{hash(doc)%10**6}")
                model.Add(weekday_cnt == sum(weekday_vars) + prior_weekday_doc)
                model.Add(festive_cnt == sum(festive_vars) + prior_festive_doc)
                model.Add(split_diff >= weekday_cnt - festive_cnt)
                model.Add(split_diff >= festive_cnt - weekday_cnt)
                extra_obj.append(type_split_penalty * split_diff)

        # Festive burden coupling: festive J and festive day duties (D/E, H/I)
        # should be distributed as alternatives, not stacked on the same doctor
        # while another eligible doctor gets none of either.
        festive_day_slots = [s for s in slots if s.rule_tag in ("Festivo_DE", "Festivo_HI")]
        festive_coupling_penalty = int(rJ.get("festive_j_day_overlap_penalty", 80_000_000) or 0)
        festive_total_spread_penalty = int(rJ.get("festive_total_spread_penalty", 35_000_000) or 0)
        festive_same_type_repeat_penalty = int(rJ.get("festive_same_type_repeat_penalty", 45_000_000) or 0)
        festive_day_consecutive_penalty = int(rJ.get("festive_day_consecutive_penalty", 70_000_000) or 0)
        heavy_quota_penalty = int(rJ.get("heavy_festive_quota_penalty", 300_000_000) or 0)
        heavy_concentration_penalty = int(rJ.get("heavy_festive_concentration_penalty", 120_000_000) or 0)
        heavy_consecutive_penalty = int(rJ.get("heavy_festive_consecutive_penalty", 220_000_000) or 0)
        heavy_unused_penalty = int(rJ.get("heavy_festive_unused_candidate_penalty", 450_000_000) or 0)
        if festive_day_slots and (
            festive_coupling_penalty > 0
            or festive_total_spread_penalty > 0
            or festive_same_type_repeat_penalty > 0
            or festive_day_consecutive_penalty > 0
            or heavy_quota_penalty > 0
            or heavy_concentration_penalty > 0
            or heavy_consecutive_penalty > 0
            or heavy_unused_penalty > 0
        ):
            festive_docs = sorted({
                doc
                for s in festive_day_slots
                for doc in (s.allowed or [])
                if doc in doctors and doc != "Recupero"
            } | set(weekend_docs))
            festive_load_vars = []
            festive_load_by_doc = {}
            festive_used_by_doc = {}
            max_prior_festive_load = 0
            for doc in festive_docs:
                j_vars = []
                for day in days:
                    if _day_is_j_festive(day):
                        v = night_var_by_day_doc.get((day.date, doc))
                        if v is not None:
                            j_vars.append(v)
                day_vars = [
                    x[(s.slot_id, doc)]
                    for s in festive_day_slots
                    if (s.slot_id, doc) in x
                ]
                prior_festive_j = _prior_j_weekend_count(doc)
                prior_festive_day = _prior_count(doc, "Festivi")
                if not j_vars and not day_vars and (prior_festive_j + prior_festive_day) <= 0:
                    continue

                j_cnt = model.NewIntVar(
                    0,
                    len(j_vars) + prior_festive_j,
                    f"fest_j_cnt_{hash(doc)%10**6}",
                )
                day_cnt = model.NewIntVar(
                    0,
                    len(day_vars) + prior_festive_day,
                    f"fest_day_cnt_{hash(doc)%10**6}",
                )
                model.Add(j_cnt == (sum(j_vars) if j_vars else 0) + prior_festive_j)
                model.Add(day_cnt == (sum(day_vars) if day_vars else 0) + prior_festive_day)

                load_upper = len(j_vars) + len(day_vars) + prior_festive_j + prior_festive_day
                load = model.NewIntVar(0, max(load_upper, 1), f"fest_total_load_{hash(doc)%10**6}")
                model.Add(load == j_cnt + day_cnt)
                festive_load_vars.append(load)
                festive_load_by_doc[doc] = load
                max_prior_festive_load = max(max_prior_festive_load, prior_festive_j + prior_festive_day)

                if heavy_unused_penalty > 0:
                    used = model.NewBoolVar(f"heavy_fest_used_{hash(doc)%10**6}")
                    unused = model.NewBoolVar(f"heavy_fest_unused_{hash(doc)%10**6}")
                    model.Add(load >= 1).OnlyEnforceIf(used)
                    model.Add(load == 0).OnlyEnforceIf(used.Not())
                    model.Add(used + unused == 1)
                    festive_used_by_doc[doc] = used
                    term = heavy_unused_penalty * unused
                    extra_obj.append(term)
                    heavy_priority_terms.append(term)

                if heavy_concentration_penalty > 0:
                    for threshold in range(2, max(load_upper, 1) + 1):
                        over_threshold = model.NewIntVar(
                            0,
                            max(load_upper, 1),
                            f"heavy_fest_conc_{threshold}_{hash(doc)%10**6}",
                        )
                        model.Add(over_threshold >= load - (threshold - 1))
                        model.Add(over_threshold >= 0)
                        term = heavy_concentration_penalty * threshold * over_threshold
                        extra_obj.append(term)
                        heavy_priority_terms.append(term)

                if festive_coupling_penalty > 0 and (j_vars or prior_festive_j > 0) and (day_vars or prior_festive_day > 0):
                    has_j = model.NewBoolVar(f"has_fest_j_{hash(doc)%10**6}")
                    has_day = model.NewBoolVar(f"has_fest_day_{hash(doc)%10**6}")
                    both = model.NewBoolVar(f"has_fest_both_{hash(doc)%10**6}")
                    model.Add(j_cnt >= 1).OnlyEnforceIf(has_j)
                    model.Add(j_cnt == 0).OnlyEnforceIf(has_j.Not())
                    model.Add(day_cnt >= 1).OnlyEnforceIf(has_day)
                    model.Add(day_cnt == 0).OnlyEnforceIf(has_day.Not())
                    model.Add(both <= has_j)
                    model.Add(both <= has_day)
                    model.Add(both >= has_j + has_day - 1)
                    extra_obj.append(festive_coupling_penalty * both)

                if festive_same_type_repeat_penalty > 0:
                    for tag in ("Festivo_DE", "Festivo_HI"):
                        tag_vars = [
                            x[(s.slot_id, doc)]
                            for s in festive_day_slots
                            if s.rule_tag == tag and (s.slot_id, doc) in x
                        ]
                        if len(tag_vars) < 2:
                            continue
                        tag_cnt = model.NewIntVar(0, len(tag_vars), f"fest_{tag}_cnt_{hash(doc)%10**6}")
                        tag_repeat = model.NewIntVar(0, len(tag_vars), f"fest_{tag}_repeat_{hash(doc)%10**6}")
                        model.Add(tag_cnt == sum(tag_vars))
                        model.Add(tag_repeat >= tag_cnt - 1)
                        model.Add(tag_repeat >= 0)
                        extra_obj.append(festive_same_type_repeat_penalty * tag_repeat)

                if festive_day_consecutive_penalty > 0:
                    festive_dates = sorted({s.day.date for s in festive_day_slots})
                    day_flags = []
                    for fdate in festive_dates:
                        date_vars = [
                            x[(s.slot_id, doc)]
                            for s in festive_day_slots
                            if s.day.date == fdate and (s.slot_id, doc) in x
                        ]
                        if not date_vars:
                            continue
                        has_festive_day = model.NewBoolVar(f"has_fest_day_{fdate}_{hash(doc)%10**6}")
                        model.AddMaxEquality(has_festive_day, date_vars)
                        day_flags.append((fdate, has_festive_day))
                    for (_date_a, flag_a), (_date_b, flag_b) in zip(day_flags, day_flags[1:]):
                        consecutive = model.NewBoolVar(
                            f"fest_day_consec_{_date_a}_{_date_b}_{hash(doc)%10**6}"
                        )
                        model.Add(consecutive <= flag_a)
                        model.Add(consecutive <= flag_b)
                        model.Add(consecutive >= flag_a + flag_b - 1)
                        extra_obj.append(festive_day_consecutive_penalty * consecutive)

                if heavy_consecutive_penalty > 0:
                    heavy_dates = sorted(
                        {s.day.date for s in festive_day_slots}
                        | {day.date for day in days if _day_is_j_festive(day)}
                    )
                    heavy_flags = []
                    for hdate in heavy_dates:
                        date_vars = [
                            x[(s.slot_id, doc)]
                            for s in festive_day_slots
                            if s.day.date == hdate and (s.slot_id, doc) in x
                        ]
                        jv = night_var_by_day_doc.get((hdate, doc))
                        if jv is not None:
                            date_vars.append(jv)
                        if not date_vars:
                            continue
                        has_heavy = model.NewBoolVar(f"has_heavy_fest_{hdate}_{hash(doc)%10**6}")
                        model.AddMaxEquality(has_heavy, date_vars)
                        heavy_flags.append((hdate, has_heavy))
                    for (_date_a, flag_a), (_date_b, flag_b) in zip(heavy_flags, heavy_flags[1:]):
                        consecutive = model.NewBoolVar(
                            f"heavy_fest_consec_{_date_a}_{_date_b}_{hash(doc)%10**6}"
                        )
                        model.Add(consecutive <= flag_a)
                        model.Add(consecutive <= flag_b)
                        model.Add(consecutive >= flag_a + flag_b - 1)
                        term = heavy_consecutive_penalty * consecutive
                        extra_obj.append(term)
                        heavy_priority_terms.append(term)

            if festive_total_spread_penalty > 0 and len(festive_load_vars) > 1:
                load_upper = len(festive_day_slots) + total_we_nights + max_prior_festive_load
                fest_total_max = model.NewIntVar(0, max(load_upper, 1), "fest_total_max")
                fest_total_min = model.NewIntVar(0, max(load_upper, 1), "fest_total_min")
                fest_total_diff = model.NewIntVar(0, max(load_upper, 1), "fest_total_diff")
                model.AddMaxEquality(fest_total_max, festive_load_vars)
                model.AddMinEquality(fest_total_min, festive_load_vars)
                model.Add(fest_total_diff == fest_total_max - fest_total_min)
                extra_obj.append(festive_total_spread_penalty * fest_total_diff)
            if heavy_quota_penalty > 0 and festive_load_by_doc:
                heavy_prior_total = sum(
                    _prior_j_weekend_count(doc) + _prior_count(doc, "Festivi")
                    for doc in festive_load_by_doc
                )
                heavy_total = len(festive_day_slots) + total_we_nights + heavy_prior_total
                n_heavy_docs = len(festive_load_by_doc)
                min_per = heavy_total // n_heavy_docs
                remainder = heavy_total - min_per * n_heavy_docs
                max_per = min_per + (1 if remainder > 0 else 0)
                for doc, load in festive_load_by_doc.items():
                    prior_load = _prior_j_weekend_count(doc) + _prior_count(doc, "Festivi")
                    cur_min = max(0, min_per - prior_load)
                    cur_max = max(0, max_per - prior_load)
                    # `load` already includes prior_load, so compare against period totals.
                    if min_per > 0:
                        under = model.NewIntVar(0, heavy_total, f"heavy_fest_under_{hash(doc)%10**6}")
                        model.Add(under >= min_per - load)
                        model.Add(under >= 0)
                        term = heavy_quota_penalty * under
                        extra_obj.append(term)
                        heavy_priority_terms.append(term)
                    over = model.NewIntVar(0, heavy_total, f"heavy_fest_over_{hash(doc)%10**6}")
                    model.Add(over >= load - max_per)
                    model.Add(over >= 0)
                    term = heavy_quota_penalty * over
                    extra_obj.append(term)
                    heavy_priority_terms.append(term)
        # Sunday J balance: Sundays are the weekend nights users inspect most.
        # Keep this separate from the generic Sat/Sun spread so a doctor does
        # not get repeated Sundays while another eligible doctor remains at 0.
        sunday_balance_penalty = int(rJ.get("sunday_night_balance_penalty", 40_000_000) or 0)
        if sunday_balance_penalty > 0:
            sunday_cnt_vars = []
            max_prior_sun = max((_prior_j_sunday_count(doc) for doc in weekend_docs), default=0)
            sunday_upper = sum(1 for day in days if day.dow == "Sun") + max_prior_sun
            for doc in sorted(weekend_docs):
                sun_vars = []
                for day in days:
                    if day.dow == "Sun":
                        v = night_var_by_day_doc.get((day.date, doc))
                        if v is not None:
                            sun_vars.append(v)
                if not sun_vars:
                    continue
                prior_sun_doc = _prior_j_sunday_count(doc)
                sun_cnt = model.NewIntVar(0, len(sun_vars) + prior_sun_doc, f"sun_j_{hash(doc)%10**6}")
                model.Add(sun_cnt == sum(sun_vars) + prior_sun_doc)
                sunday_cnt_vars.append(sun_cnt)
            if len(sunday_cnt_vars) > 1:
                sun_max = model.NewIntVar(0, max(sunday_upper, 1), "sun_j_max")
                sun_min = model.NewIntVar(0, max(sunday_upper, 1), "sun_j_min")
                sun_diff = model.NewIntVar(0, max(sunday_upper, 1), "sun_j_diff")
                model.AddMaxEquality(sun_max, sunday_cnt_vars)
                model.AddMinEquality(sun_min, sunday_cnt_vars)
                model.Add(sun_diff == sun_max - sun_min)
                extra_obj.append(sunday_balance_penalty * sun_diff)
    # H monthly quotas Mon-Fri
    # MODIFICA 1: Grimaldi e Calabrò sono esclusi da H; ignora eventuali quote riferite a loro
    _h_df_pair = {norm_name("Grimaldi"), norm_name("Calabrò")}
    if "rules" in cfg and "H" in cfg["rules"]:
        mqH = cfg["rules"]["H"].get("monthly_quotas") or {}
        for key, q in mqH.items():
            # keys could be 'Grimaldi_mon_fri'
            m = re.match(r"(.+)_mon_fri", str(key).strip(), flags=re.I)
            if not m:
                continue
            doc = norm_name(m.group(1))
            if doc not in doctors:
                continue
            if doc in _h_df_pair:
                continue   # Grimaldi/Calabrò non vanno mai in H
            vars_ = []
            for day in days:
                if day.dow in ["Mon","Tue","Wed","Thu","Fri"]:
                    sH = next((s for s in slots_by_day[day.date] if s.columns == ["H"]), None)
                    if sH and (sH.slot_id, doc) in x:
                        vars_.append(x[(sH.slot_id, doc)])
            if vars_:
                q_target = min(_monthly_target_for_period(int(q), doc, "H"), len(vars_))
                q_cap = min(_monthly_cap_for_period(int(q), doc, "H"), len(vars_))
                q_cap = max(q_cap, q_target)
                model.Add(sum(vars_) <= q_cap)
                _hdef = model.NewIntVar(0, len(vars_), f"hdef_{hash((doc,key))%10**6}")
                model.Add(_hdef >= q_target - sum(vars_))
                model.Add(_hdef >= 0)
                extra_obj.append(J_QUOTA_DEV_PENALTY * _hdef)
                if q_cap > q_target:
                    _hsup = model.NewIntVar(0, len(vars_), f"hsup_{hash((doc,key))%10**6}")
                    model.Add(_hsup >= sum(vars_) - q_target)
                    model.Add(_hsup >= 0)
                    extra_obj.append(J_QUOTA_DEV_PENALTY * _hsup)
        # cap per doctor for pool_mon_fri
        cap = cfg["rules"]["H"].get("cap_mon_fri_per_doctor")
        if cap is not None:
            cap = int(cap)
            pool_cap = [norm_name(d) for d in (cfg["rules"]["H"].get("pool_mon_fri") or [])]
            for doc in pool_cap:
                if doc not in doctors:
                    continue
                vars_ = []
                for day in days:
                    if day.dow in ["Mon","Tue","Wed","Thu","Fri"]:
                        sH = next((s for s in slots_by_day[day.date] if s.columns == ["H"]), None)
                        if sH and (sH.slot_id, doc) in x:
                            vars_.append(x[(sH.slot_id, doc)])
                if vars_:
                    # Soft: penalizza eccesso rispetto al cap (no hard per evitare infeasibility)
                    _hcap_over = model.NewIntVar(0, len(vars_), f"hcap_over_{hash(doc)%10**6}")
                    model.Add(_hcap_over >= sum(vars_) - cap)
                    model.Add(_hcap_over >= 0)
                    extra_obj.append(500_000 * _hcap_over)
    # L quota Recupero
    if "rules" in cfg and "L" in cfg["rules"]:
        qrec = cfg["rules"]["L"].get("quota_recupero_per_month")
        if qrec is not None and "Recupero" in doctors:
            vars_=[]
            for s in slots:
                if s.columns == ["L"] and (s.slot_id, "Recupero") in x:
                    vars_.append(x[(s.slot_id, "Recupero")])
            if vars_:
                qrec_int = _monthly_target_for_period(int(qrec), "Recupero", "L")
                model.Add(sum(vars_) <= qrec_int)
                # Soft: penalizza shortfall senza hard equality (evita infeasibility)
                _lrec_sum = model.NewIntVar(0, qrec_int, f"L_rec_sum")
                model.Add(_lrec_sum == sum(vars_))
                _lrec_short = model.NewIntVar(0, qrec_int, f"L_rec_short")
                model.Add(_lrec_short >= qrec_int - _lrec_sum)
                model.Add(_lrec_short >= 0)
                extra_obj.append(5 * _lrec_short)

    
    # ------------------------------------------------------------
    # Vincoli richiesti aggiuntivi (Roberto)
    # ------------------------------------------------------------
    # Y (Ambulatori) – due lunedì/mese: Recupero deve risultare come affiancamento,
    # ma SENZA creare uno slot extra (per evitare di "consumare" un medico in più).
    # Implementazione: imponiamo che in esattamente 2 lunedì/mese T=Recupero;
    # in output, quando T=Recupero di lunedì, aggiungiamo "Recupero" in Y come seconda riga.
    if "rules" in cfg and "Y" in cfg["rules"] and "T" in cfg["rules"]:
        rY = cfg["rules"]["Y"] or {}
        if (
            rY.get("recupero_two_mondays_per_month", False)
            and rY.get("recupero_affianca_in_T", False)
            and "Recupero" in doctors
        ):
            # Fixed choice: first 2 Mondays of the month
            monday_days = [d for d in days if d.dow == "Mon"]
            monday_days.sort(key=lambda d: d.date)
            target_mondays = monday_days[:2]
            other_mondays = monday_days[2:]

            # Save for diagnostics/logging
            MonT_need_rec = len(target_mondays)
            MonT_target_dates = [d.date for d in target_mondays]

            # Soft: forte preferenza T=Recupero sui primi 2 lunedì (no hard per evitare infeasibility)
            _REC_T_MON_PENALTY = 6_000_000
            for d in target_mondays:
                sT = next((s for s in slots_by_day[d.date] if s.columns == ["T"]), None)
                if sT and (sT.slot_id, "Recupero") in x:
                    _v_rec_t = x[(sT.slot_id, "Recupero")]
                    _b_rec_t = model.NewBoolVar(f"rec_t_miss_{d.date.isoformat()}")
                    model.Add(_b_rec_t + _v_rec_t >= 1)  # almeno uno dei due è 1
                    extra_obj.append(_REC_T_MON_PENALTY * _b_rec_t)

            # Soft: scoraggia T=Recupero sugli altri lunedì
            for d in other_mondays:
                sT = next((s for s in slots_by_day[d.date] if s.columns == ["T"]), None)
                if sT and (sT.slot_id, "Recupero") in x:
                    extra_obj.append((_REC_T_MON_PENALTY // 2) * x[(sT.slot_id, "Recupero")])

    # U (Contr.PM) – Cimino esattamente N volte/mese (default: 2)
    if "rules" in cfg and "U" in cfg["rules"] and "Cimino" in doctors:
        rU = cfg["rules"]["U"] or {}
        exact = int(rU.get("cimino_exact_per_month", 0) or 0)
        if exact > 0:
            vars_ = []
            for s in slots:
                if s.columns == ["U"] and (s.slot_id, "Cimino") in x:
                    vars_.append(x[(s.slot_id, "Cimino")])
            if vars_:
                effective_exact = min(_monthly_target_for_period(exact, "Cimino", "U"), len(vars_))
                effective_cap = min(_monthly_cap_for_period(exact, "Cimino", "U"), len(vars_))
                effective_cap = max(effective_cap, effective_exact)
                # Hard upper bound + soft lower
                model.Add(sum(vars_) <= effective_cap)
                _cudef = model.NewIntVar(0, len(vars_), f"cudef_{effective_exact}")
                model.Add(_cudef >= effective_exact - sum(vars_))
                model.Add(_cudef >= 0)
                extra_obj.append(J_QUOTA_DEV_PENALTY * _cudef)
                if effective_cap > effective_exact:
                    _cusup = model.NewIntVar(0, len(vars_), f"cusup_{effective_exact}")
                    model.Add(_cusup >= sum(vars_) - effective_exact)
                    model.Add(_cusup >= 0)
                    extra_obj.append(J_QUOTA_DEV_PENALTY * _cusup)
                if effective_exact < exact and not period_is_partial_month:
                    pre_solve_warnings.append(
                        f"Cimino U: richiesti esattamente {exact} ma solo {len(vars_)} slot disponibili, ridotto a {effective_exact}."
                    )
            else:
                pre_solve_warnings.append(
                    f"Cimino non ha slot U disponibili questo mese: vincolo cimino_exact_per_month ignorato"
                )

    # MODIFICA 3: se lunedì V=Allegra allora U deve essere Crea o Dattilo (hard constraint nel solver)
    if "rules" in cfg and "U" in cfg["rules"] and "V" in cfg["rules"]:
        rU_c = cfg["rules"]["U"] or {}
        if rU_c.get("v_allegra_monday_constraint", False):
            _allegra = norm_name("Allegra")
            _crea = norm_name("Crea")
            _dattilo = norm_name("Dattilo")
            _forbidden_in_u_if_v_allegra = [d for d in doctors if norm_name(d) not in {_crea, _dattilo, _allegra}]
            for day in [d for d in days if d.dow == "Mon"]:
                sV = next((s for s in slots_by_day[day.date] if s.columns == ["V"]), None)
                sU = next((s for s in slots_by_day[day.date] if s.columns == ["U"]), None)
                if sV is None or sU is None:
                    continue
                v_allegra = x.get((sV.slot_id, _allegra))
                if v_allegra is None:
                    continue
                # Se V=Allegra il lunedì → U deve essere Crea o Dattilo
                for forb in _forbidden_in_u_if_v_allegra:
                    u_forb = x.get((sU.slot_id, forb))
                    if u_forb is not None:
                        # u_forb=0 when v_allegra=1
                        model.Add(u_forb == 0).OnlyEnforceIf(v_allegra)

    # I (Cardiologia pomeriggio) – De Gregorio max N nei feriali (i festivi sono HI, quindi esclusi)
    if "rules" in cfg and "I" in cfg["rules"] and "De Gregorio" in doctors:
        rI = cfg["rules"]["I"] or {}
        max_i = int(rI.get("degregorio_max_weekdays", 0) or 0)
        if max_i > 0:
            vars_ = []
            for s in slots:
                # Count only real I slots on Mon-Sat. Festivi use the unified HI slot.
                if (
                    s.columns == ["I"]
                    and getattr(s.day, "dow", "") in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
                    and (s.slot_id, "De Gregorio") in x
                ):
                    vars_.append(x[(s.slot_id, "De Gregorio")])
            if vars_:
                model.Add(sum(vars_) <= max_i)
# ------------------------------------------------------------
    # Vincoli specifici per 'Recupero' (richiesti):
    # - T (Interni): almeno N/mese (hard) + target soft (es. 5)
    # - Q (ECO base): massimo N/mese
    # ------------------------------------------------------------
    if "rules" in cfg and "T" in cfg["rules"]:
        rT = cfg["rules"]["T"] or {}
        if "Recupero" in doctors:
            min_rec = _monthly_target_for_period(int(rT.get("recupero_min_per_month", 0) or 0), "Recupero", "T")
            target_rec = _monthly_target_for_period(int(rT.get("recupero_target_per_month", 0) or 0), "Recupero", "T")
            target_pen = int(rT.get("recupero_target_penalty", 2000) or 2000)
            vars_ = []
            for s in slots:
                if s.columns == ["T"] and (s.slot_id, "Recupero") in x:
                    vars_.append(x[(s.slot_id, "Recupero")])
            if vars_:
                cnt = model.NewIntVar(0, len(vars_), "T_rec_cnt")
                model.Add(cnt == sum(vars_))
                if min_rec > 0:
                    effective_min_rec = min(min_rec, len(vars_))
                    # Soft: fortissima penalità per mancanza (non hard per evitare infeasibility)
                    _trecdef = model.NewIntVar(0, effective_min_rec, f"trecdef")
                    model.Add(_trecdef >= effective_min_rec - cnt)
                    model.Add(_trecdef >= 0)
                    extra_obj.append(J_QUOTA_DEV_PENALTY * _trecdef)
                    if effective_min_rec < min_rec:
                        pre_solve_warnings.append(
                            f"Recupero T min: richiesto min {min_rec} ma solo {len(vars_)} slot T disponibili, ridotto a {effective_min_rec}."
                        )
                if target_rec > 0:
                    short = model.NewIntVar(0, target_rec, "T_rec_target_short")
                    model.Add(cnt + short >= target_rec)
                    extra_obj.append(target_pen * short)

    if "rules" in cfg and "Q" in cfg["rules"]:
        rQ = cfg["rules"]["Q"] or {}
        if "Recupero" in doctors:
            max_rec = _monthly_cap_for_period(int(rQ.get("recupero_max_per_month", 0) or 0), "Recupero", "Q")
            if max_rec > 0:
                vars_ = []
                for s in slots:
                    if s.columns == ["Q"] and (s.slot_id, "Recupero") in x:
                        vars_.append(x[(s.slot_id, "Recupero")])
                if vars_:
                    model.Add(sum(vars_) <= max_rec)
        # Q: hard cap per ogni medico del pool per evitare dominanza
        # Ideale: round(n_Q_slots / pool_size) + 1
        q_pool_raw = [norm_name(d) for d in (rQ.get("pool") or []) if norm_name(d) in doctors]
        q_slots = [s for s in slots if s.columns == ["Q"]]
        if q_pool_raw and q_slots:
            import math as _math
            q_cap = _math.ceil(len(q_slots) / len(q_pool_raw)) + 1
            for doc in q_pool_raw:
                vars_ = [x[(s.slot_id, doc)] for s in q_slots if (s.slot_id, doc) in x]
                if vars_:
                    model.Add(sum(vars_) <= q_cap)

    # T: hard cap per ogni medico del pool per evitare dominanza (Cimino/D'Angelo/Recupero a 5)
    if "rules" in cfg and "T" in cfg["rules"]:
        rT2 = cfg["rules"]["T"] or {}
        t_pool_raw = [norm_name(d) for d in (rT2.get("pool") or []) if norm_name(d) in doctors]
        t_slots = [s for s in slots if s.columns == ["T"]]
        if t_pool_raw and t_slots:
            import math as _math
            t_cap = _math.ceil(len(t_slots) / len(t_pool_raw)) + 1
            for doc in t_pool_raw:
                vars_ = [x[(s.slot_id, doc)] for s in t_slots if (s.slot_id, doc) in x]
                if vars_:
                    model.Add(sum(vars_) <= t_cap)

    # W: cap Recupero per evitare che domini il turno
    if "rules" in cfg and "W" in cfg["rules"] and "Recupero" in doctors:
        rW = cfg["rules"]["W"] or {}
        w_rec_max = _monthly_cap_for_period(int(rW.get("recupero_max_per_month", 0) or 0), "Recupero", "W")
        if w_rec_max > 0:
            vars_ = []
            for s in slots:
                if s.columns == ["W"] and (s.slot_id, "Recupero") in x:
                    vars_.append(x[(s.slot_id, "Recupero")])
            if vars_:
                model.Add(sum(vars_) <= w_rec_max)

    # ---------------------------------------------------------------
    # VINCOLI UNIVERSITARI — calcolati PRIMA delle assegnazioni generali
    # Zito, Dattilo, De Gregorio: monte ore = 60% degli ospedalieri.
    #
    #   working_days = giorni lun-sab del mese (es. marzo 2026 = 26)
    #   target = round(working_days × 0.6)  → es. 16
    #
    #   Pesi:
    #   - J (notte 12h) = 2 turni per Zito e Dattilo
    #   - Ogni altro turno operativo = 1 (incluso V)
    #   - C (Reperibilità) = NON conta (turno di guardia extra, non monte ore)
    #
    #   Vincoli:
    #   - CAP HARD (≤ target+1): non si può sforare il contratto
    #   - FLOOR SOFT (penalità se < target): il solver cerca di avvicinarsi
    #     senza rendere infeasible se il pool è ristretto (es. Zito solo Q/R/S/J)
    # ---------------------------------------------------------------
    UNIV_EXCLUDE_COLS = {"C"}   # solo Reperibilità esclusa
    gc_uni = gc.get("university_doctors") or {}
    uni_ratio = float(gc.get("university_ratio", 0.6))
    if gc_uni and uni_ratio > 0:
        working_days = sum(1 for d in days if d.dow in ["Mon","Tue","Wed","Thu","Fri","Sat"] and not is_festivo(d, cfg))
        target = round(working_days * uni_ratio)
        for doc_raw, doc_cfg in gc_uni.items():
            doc = norm_name(doc_raw)
            if doc not in doctors:
                continue
            night_double = bool((doc_cfg or {}).get("night_counts_double", False))
            counts_as_map = cfg.get("pool_counts_as") or {}
            weighted_terms = []
            for s in slots:
                v = x.get((s.slot_id, doc))
                if v is None:
                    continue
                col_letters = s.columns or []
                if counts_as_map:
                    # counts_as da pool_config: C=0 (non conta), J=2 (vale doppio), altri=1
                    weight = max((counts_as_map.get(c, 1) for c in col_letters), default=1)
                else:
                    # Retrocompatibilità: escludi C, usa night_counts_double per J
                    if any(c in UNIV_EXCLUDE_COLS for c in col_letters):
                        weight = 0
                    else:
                        is_night = "J" in col_letters
                        weight = 2 if (night_double and is_night) else 1
                if weight == 0:
                    continue
                weighted_terms.append(weight * v)
            if not weighted_terms:
                continue
            max_possible = len(weighted_terms) * 2
            uni_cnt = model.NewIntVar(0, max_possible, f"uni_cnt_{hash(doc)%10**6}")
            model.Add(uni_cnt == sum(weighted_terms))
            # Soft: penalizza se sopra target+1 (no hard per evitare infeasibility combinatoria)
            _uni_over = model.NewIntVar(0, max_possible, f"uni_over_{hash(doc)%10**6}")
            model.Add(_uni_over >= uni_cnt - (target + 1))
            model.Add(_uni_over >= 0)
            extra_obj.append(500_000 * _uni_over)
            # FLOOR SOFT: penalizza se sotto target
            under = model.NewIntVar(0, target, f"uni_under_{hash(doc)%10**6}")
            model.Add(under >= target - uni_cnt)
            model.Add(under >= 0)
            extra_obj.append(500 * under)

    # AB: MODIFICA 5 — giovedì BILANCIATO (nessuna preferenza per Crea), sabati HARD con Crea
    if "rules" in cfg and "AB" in cfg["rules"]:
        rAB = cfg["rules"]["AB"]
        # Giovedì: soft balance tra tutti i medici del pool AB
        ab_thu_slots = [s for s in slots if s.rule_tag == "AB"]
        ab_thu_pool = []
        _seen = set()
        for _d in (rAB.get("fallback_pool") or []):
            _dn = norm_name(_d)
            if _dn not in _seen and _dn in doctors:
                ab_thu_pool.append(_dn)
                _seen.add(_dn)
        if ab_thu_slots and ab_thu_pool:
            n_thu = len(ab_thu_slots)
            ab_max_per_doc = model.NewIntVar(0, n_thu, "AB_thu_max")
            for doc in ab_thu_pool:
                vars_ = [x[(s.slot_id, doc)] for s in ab_thu_slots if (s.slot_id, doc) in x]
                if vars_:
                    cnt_doc = model.NewIntVar(0, n_thu, f"AB_thu_cnt_{hash(doc)%10**6}")
                    model.Add(cnt_doc == sum(vars_))
                    model.Add(ab_max_per_doc >= cnt_doc)
            extra_obj.append(50 * ab_max_per_doc)  # minimizza il massimo → bilanciamento

        # Sabati: HARD con Crea (2 sabati/mese)
        sat_doc = norm_name(rAB.get("saturday_only_doctor") or "Crea")
        sat_n = _monthly_target_for_period(int(rAB.get("saturday_per_month", 0) or 0), sat_doc, "AB")
        sat_cap = _monthly_cap_for_period(int(rAB.get("saturday_per_month", 0) or 0), sat_doc, "AB")
        if sat_n > 0 and sat_doc in doctors:
            vars_=[]
            for s in slots:
                if s.rule_tag == "AB_SAT" and (s.slot_id, sat_doc) in x:
                    vars_.append(x[(s.slot_id, sat_doc)])
            if vars_:
                if bool(rAB.get("saturday_soft", False)):
                    # soft (fallback per evitare infeasible)
                    model.Add(sum(vars_) <= sat_cap)
                    short = model.NewIntVar(0, sat_n, "AB_sat_short")
                    model.Add(sum(vars_) + short == sat_n)
                    extra_obj.append(int(rAB.get("saturday_shortfall_penalty", 10000)) * short)
                else:
                    # HARD: i 2 sabati devono essere SOLO con Crea
                    model.Add(sum(vars_) == sat_n)

    # Weekend full off: at least N Sat+Sun "full weekends off" per doctor.
    # By default this is a HARD constraint. If it makes the month infeasible,
    # you can set global_constraints.weekend_off_soft: true to make it a SOFT constraint
    # (the solver will minimize the number of missing weekends-off).
    min_weekends = _monthly_target_for_period(int(gc.get("min_full_weekends_off_per_month", 0) or 0))
    weekend_exempt = set(norm_name(x) for x in (gc.get("weekend_off_exempt") or []))
    # NOTA: 'Recupero' è trattato come medico reale: rientra nel conteggio dei weekend-off.
    weekend_soft = bool(gc.get("weekend_off_soft", False))
    weekend_penalty = int(gc.get("weekend_off_penalty", 50) or 50)  # penalty per missing full weekend off
    weekend_shortfalls: List = []
    if min_weekends > 0:
        # find weekend pairs within available days
        date_set = {d.date for d in days}
        weekend_pairs = []
        for drow in days:
            if drow.dow == "Sat":
                sun = drow.date + dt.timedelta(days=1)
                # count only complete Sat+Sun pairs present in the template for this month
                if sun in date_set and DOW_MAP[sun.weekday()] == "Sun":
                    weekend_pairs.append((drow.date, sun))
        # For each doctor, create weekend_off[w,doc] bool
        weekend_off: Dict[Tuple[int, str], "cp_model.IntVar"] = {}
        for wi, (sat, sun) in enumerate(weekend_pairs):
            for doc in doctors:
                if doc in weekend_exempt:
                    continue
                b = model.NewBoolVar(f"wkoff_{wi}_{hash(doc)%10**6}")
                weekend_off[(wi, doc)] = b
                # If weekend_off=1 then doc has no assignment on sat and sun
                # NOTE: C (Reperibilità) is excluded — it's passive on-call,
                # not an active shift, so it doesn't consume a weekend off.
                for s in slots_by_day.get(sat, []):
                    if s.columns == ["C"]:
                        continue
                    if (s.slot_id, doc) in x:
                        model.Add(x[(s.slot_id, doc)] == 0).OnlyEnforceIf(b)
                for s in slots_by_day.get(sun, []):
                    if s.columns == ["C"]:
                        continue
                    if (s.slot_id, doc) in x:
                        model.Add(x[(s.slot_id, doc)] == 0).OnlyEnforceIf(b)
                # Reverse implication: if any assignment on sat or sun then b=0
                any_vars = []
                for s in slots_by_day.get(sat, []):
                    if s.columns == ["C"]:
                        continue
                    v = x.get((s.slot_id, doc))
                    if v is not None:
                        any_vars.append(v)
                for s in slots_by_day.get(sun, []):
                    if s.columns == ["C"]:
                        continue
                    v = x.get((s.slot_id, doc))
                    if v is not None:
                        any_vars.append(v)
                if any_vars:
                    # sum(any_vars)==0 -> b=1
                    model.Add(sum(any_vars) == 0).OnlyEnforceIf(b)
                    # if any assignment then b=0
                    for av in any_vars:
                        model.Add(b + av <= 1)
        for doc in doctors:
            if doc in weekend_exempt:
                continue
            vars_ = [weekend_off[(wi, doc)] for wi, _ in enumerate(weekend_pairs) if (wi, doc) in weekend_off]
            if vars_:
                if weekend_soft:
                    short = model.NewIntVar(0, len(weekend_pairs), f"wkshort_{hash(doc)%10**6}")
                    model.Add(sum(vars_) + short >= min_weekends)
                    weekend_shortfalls.append(short)
                else:
                    model.Add(sum(vars_) >= min_weekends)
    # Objectives (soft): fairness + maximize S dedicated + minimize weekend night concentration + prefer night spacing >=7
    objective_terms = []
    # Weekend-off soft penalties
    if weekend_shortfalls:
        objective_terms.append(weekend_penalty * sum(weekend_shortfalls))
    # fairness: minimize max assignments per doctor
    real_doctors = list(doctors)
    max_load = model.NewIntVar(0, 999, "max_load")
    for doc in real_doctors:
        vars_ = []
        for s in slots:
            v = x.get((s.slot_id, doc))
            if v is not None:
                vars_.append(v)
        if vars_:
            load = model.NewIntVar(0, 999, f"load_{hash(doc)%10**6}")
            model.Add(load == sum(vars_) + _prior_total_count(doc))
            model.Add(load <= max_load)
    objective_terms.append(max_load * 10)
    # Column-specific balancing (and optional soft caps) for specific columns.
    # If a rule has `balance: true`, or defines a `distribution_pool`, we try to balance assignments within that column.
    # If a rule defines `max_per_doctor: N`, we penalize assignments above N (soft cap).
    # NOTE: a hard cap is not enforced here because it can easily make the model infeasible when the pool is small;
    # the soft cap is still reported via resulting counts (and can be made hard by widening the pool).
    try:
        for col_key, rcol in (rules_map or {}).items():
            if not isinstance(rcol, dict):
                continue
            if not (rcol.get('balance') or rcol.get('distribution_pool') or rcol.get('max_per_doctor')):
                continue
            # consider only slots whose rule_tag matches this column key (avoid Festivo_HI etc.)
            slot_ids = [s.slot_id for s in slots if (getattr(s, 'rule_tag', '') == col_key)]
            if not slot_ids:
                continue
            # prefer explicit pools; otherwise infer from variables
            pool_raw = (rcol.get('pool') or rcol.get('distribution_pool') or rcol.get('allowed') or [])
            pool = []
            if isinstance(pool_raw, list):
                for p in pool_raw:
                    pn = norm_name(p)
                    if pn and pn in doc_to_idx :
                        pool.append(pn)
            if not pool:
                # fallback: any doctor that actually appears in some var for these slots
                pool = []
                for d in real_doctors:
                    for sid in slot_ids:
                        if x.get((sid, d)) is not None:
                            pool.append(d)
                            break
            if not pool:
                continue
            bal_w = int(rcol.get('balance_weight') or 40)
            cap = int(rcol.get('max_per_doctor') or 0)
            cap_pen = int(rcol.get('max_per_doctor_penalty') or 800)
            max_prior_col = max((_prior_rule_balance_count(d, col_key) for d in pool), default=0)
            max_col = model.NewIntVar(0, len(slot_ids) + max_prior_col, f"max_{col_key}_load")
            _load_vars_for_spread = []
            for d in pool:
                vars_d = []
                for sid in slot_ids:
                    v = x.get((sid, d))
                    if v is not None:
                        vars_d.append(v)
                if not vars_d:
                    continue
                prior_col = _prior_rule_balance_count(d, col_key)
                load_d = model.NewIntVar(0, len(slot_ids) + prior_col, f"load_{col_key}_{hash(d)%10**6}")
                model.Add(load_d == sum(vars_d) + prior_col)
                model.Add(load_d <= max_col)
                _load_vars_for_spread.append(load_d)
                if cap > 0:
                    over = model.NewIntVar(0, len(slot_ids) + prior_col, f"over_{col_key}_{hash(d)%10**6}")
                    # over >= load_d - cap, over >= 0
                    model.Add(load_d - cap <= over)
                    model.Add(over >= 0)
                    objective_terms.append(over * cap_pen)
            objective_terms.append(max_col * bal_w)
            # Penalizza anche lo spread max-min per forzare equidistribuzione
            if len(_load_vars_for_spread) >= 2:
                min_col = model.NewIntVar(0, len(slot_ids) + max_prior_col, f"min_{col_key}_load")
                model.AddMinEquality(min_col, _load_vars_for_spread)
                spread_col = model.NewIntVar(0, len(slot_ids) + max_prior_col, f"spread_{col_key}_load")
                model.Add(spread_col == max_col - min_col)
                objective_terms.append(spread_col * bal_w)
    except Exception:
        # never fail scheduling due to a balance/cap config issue
        pass
    # Maximize number of Wednesdays with dedicated S assignment (if optional)
    s_slots = [s for s in slots if s.columns == ["S"]]
    s_dedicated = []
    for s in s_slots:
        vars_ = [x.get((s.slot_id, d)) for d in doctors if (s.slot_id, d) in x]
        vars_ = [v for v in vars_ if v is not None]
        if vars_:
            b = model.NewBoolVar(f"sfilled_{hash(s.slot_id)%10**6}")
            model.Add(sum(vars_) == 1).OnlyEnforceIf(b)
            model.Add(sum(vars_) == 0).OnlyEnforceIf(b.Not())
            s_dedicated.append(b)
    if s_dedicated:
        # subtract to maximize (minimize negative)
        objective_terms.append(-5 * sum(s_dedicated))
    # Prefer night spacing >=7: penalize gaps of 5 or 6
    preferred = int(gc.get("night_spacing_days_preferred", 7))
    if preferred > min_gap:
        penalties=[]
        for doc in real_doctors:
            for i, drow in enumerate(days):
                for k in range(min_gap, preferred):
                    j=i+k
                    if j>=len(days): break
                    v1=night_var_by_day_doc.get((drow.date, doc))
                    v2=night_var_by_day_doc.get((days[j].date, doc))
                    if v1 is not None and v2 is not None:
                        p = model.NewBoolVar(f"ngap_{hash(doc)%10**6}_{i}_{k}")
                        # p=1 if both nights, else 0
                        model.Add(v1 + v2 == 2).OnlyEnforceIf(p)
                        model.Add(v1 + v2 <= 1).OnlyEnforceIf(p.Not())
                        penalties.append(p)
        if penalties:
            objective_terms.append(3 * sum(penalties))
    # Weekend night concentration: già gestito nel blocco J sopra con hard max=2 e strong soft

    # ── Historical balance: soft constraints from previous months ────────────
    _hist = historical_stats or {}
    if _hist:
        HIST_NIGHT_PENALTY = 150
        HIST_FEST_PENALTY = 200

        # Notti J: penalizza medici con più notti storiche
        _rJ_cfg = (cfg.get("rules") or {}).get("J") or {}
        _mq_fixed_hist = set()
        for _k in (_rJ_cfg.get("monthly_quotas") or {}).keys():
            _mq_fixed_hist.add(norm_name(_k))
        try:
            _night_pool_hist = night_pool  # defined inside "J" block
        except NameError:
            _night_pool_hist = set()
        _night_pool_free_hist = [d for d in sorted(_night_pool_hist) if d not in _mq_fixed_hist]

        for _doc in _night_pool_free_hist:
            _j_data = _hist.get(_doc, {}).get("J", {})
            _hist_j_total = _j_data.get("total", 0) if isinstance(_j_data, dict) else 0
            if _hist_j_total <= 0:
                continue
            _vars = [night_var_by_day_doc.get((d.date, _doc)) for d in days
                     if night_var_by_day_doc.get((d.date, _doc)) is not None]
            if _vars:
                _cnt = model.NewIntVar(0, len(days), f"hist_j_{abs(hash(_doc)) % 10 ** 6}")
                model.Add(_cnt == sum(_vars))
                extra_obj.append(HIST_NIGHT_PENALTY * _hist_j_total * _cnt)

        # Domeniche J: penalizza notti domenicali storiche
        for _doc in sorted(_night_pool_hist):
            _j_data2 = _hist.get(_doc, {}).get("J", {})
            _hist_dom_j = _j_data2.get("domeniche", 0) if isinstance(_j_data2, dict) else 0
            if _hist_dom_j <= 0:
                continue
            _sun_vars = [night_var_by_day_doc.get((d.date, _doc))
                         for d in days if d.dow == "Sun"
                         if night_var_by_day_doc.get((d.date, _doc)) is not None]
            if _sun_vars:
                _sun_cnt = model.NewIntVar(0, len(_sun_vars), f"hist_sunj_{abs(hash(_doc)) % 10 ** 6}")
                model.Add(_sun_cnt == sum(_sun_vars))
                extra_obj.append(HIST_FEST_PENALTY * _hist_dom_j * _sun_cnt)

        # Sabati J: penalizza notti di sabato storiche (stessa logica delle domeniche)
        for _doc in sorted(_night_pool_hist):
            _j_data3 = _hist.get(_doc, {}).get("J", {})
            _hist_sab_j = _j_data3.get("sabati", 0) if isinstance(_j_data3, dict) else 0
            if _hist_sab_j <= 0:
                continue
            _sat_vars = [night_var_by_day_doc.get((d.date, _doc))
                         for d in days if d.dow == "Sat"
                         if night_var_by_day_doc.get((d.date, _doc)) is not None]
            if _sat_vars:
                _sat_cnt = model.NewIntVar(0, len(_sat_vars), f"hist_satj_{abs(hash(_doc)) % 10 ** 6}")
                model.Add(_sat_cnt == sum(_sat_vars))
                extra_obj.append(HIST_FEST_PENALTY * _hist_sab_j * _sat_cnt)

        # Festivi DEHI: penalizza medici con più festivi storici (solo pool senza quota fissa)
        HIST_DEHI_PENALTY = 200
        _rFest_cfg = (cfg.get("rules") or {}).get("Festivi") or {}
        _fest_fixed_hist = {norm_name(k) for k in (_rFest_cfg.get("quotas") or {}).keys()}
        _fest_pool_hist = [norm_name(d) for d in (_rFest_cfg.get("pool") or [])
                          if norm_name(d) in doctors and norm_name(d) != "Recupero"]
        _fest_pool_free_hist = [d for d in _fest_pool_hist if d not in _fest_fixed_hist]
        _festivo_slots_hist = [s for s in slots if s.rule_tag in ("Festivo_DE", "Festivo_HI")]

        if _fest_pool_free_hist and _festivo_slots_hist:
            for _doc in _fest_pool_free_hist:
                _hist_dehi = _hist.get(_doc, {}).get("_festivi_DE_HI", 0)
                if not isinstance(_hist_dehi, int):
                    _hist_dehi = 0
                if _hist_dehi <= 0:
                    continue
                _fvars = [x.get((s.slot_id, _doc)) for s in _festivo_slots_hist
                          if (s.slot_id, _doc) in x]
                if _fvars:
                    _fcnt = model.NewIntVar(0, len(_festivo_slots_hist),
                                            f"hist_dehi_{abs(hash(_doc)) % 10 ** 6}")
                    model.Add(_fcnt == sum(_fvars))
                    extra_obj.append(HIST_DEHI_PENALTY * _hist_dehi * _fcnt)

    full_objective_terms = objective_terms + extra_obj + coverage_obj_terms
    coverage_expr = sum(coverage_obj_terms) if coverage_obj_terms else 0

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float((gc.get("solver_max_time_seconds", 60.0) or 60.0))
    solver.parameters.num_search_workers = 1
    coverage_status = None
    coverage_cost = None
    heavy_status = None
    heavy_cost = None

    def _status_name(st) -> Optional[str]:
        if st is None:
            return None
        try:
            return solver.StatusName(st)
        except Exception:
            return str(st)

    if coverage_obj_terms:
        model.Minimize(coverage_expr)
        coverage_status = solver.Solve(model)
        if coverage_status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            coverage_cost = int(solver.Value(coverage_expr))
            if coverage_status == cp_model.OPTIMAL:
                model.Add(coverage_expr == coverage_cost)
            else:
                # If coverage proof times out, keep the best found coverage as
                # a non-worsening ceiling. Otherwise the final objective can
                # drift back to a schedule with more blanks.
                model.Add(coverage_expr <= coverage_cost)

    heavy_priority_expr = sum(heavy_priority_terms) if heavy_priority_terms else 0
    if heavy_priority_terms:
        model.Minimize(heavy_priority_expr)
        heavy_status = solver.Solve(model)
        if heavy_status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            heavy_cost = int(solver.Value(heavy_priority_expr))
            if heavy_status == cp_model.OPTIMAL:
                model.Add(heavy_priority_expr == heavy_cost)
            else:
                model.Add(heavy_priority_expr <= heavy_cost)

    model.Minimize(sum(full_objective_terms))
    status = solver.Solve(model)
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        # ── Diagnostica: retry senza vincoli di quota hard per identificare la causa ──
        _diag_hints: List[str] = []
        try:
            _model2 = cp_model.CpModel()
            _extra2: List = []
            _BLANK_PEN2 = 5_000_000
            _BLANK_CRIT2 = 50_000_000
            _CRIT2 = {"H", "I", "J", "Festivo_HI"}
            _x2: Dict = {}
            for s2 in slots:
                for d2 in (s2.allowed or []):
                    if norm_name(d2) in doc_to_idx:
                        _x2[(s2.slot_id, norm_name(d2))] = _model2.NewBoolVar(
                            f"x2_{hash(s2.slot_id)%10**7}_{hash(d2)%10**7}"
                        )
            for s2 in slots:
                _v2 = [_x2[(s2.slot_id, norm_name(d2))] for d2 in (s2.allowed or [])
                       if (s2.slot_id, norm_name(d2)) in _x2]
                if not _v2:
                    continue
                if s2.required:
                    _bb2 = _model2.NewBoolVar(f"bb2_{hash(s2.slot_id)%10**7}")
                    _model2.Add(sum(_v2) + _bb2 == 1)
                    _pen2 = _BLANK_CRIT2 if s2.rule_tag in _CRIT2 else _BLANK_PEN2
                    _extra2.append(_pen2 * _bb2)
                else:
                    _model2.Add(sum(_v2) <= 1)
            # one_per_day (semplificato, senza share slack)
            for _d2 in days:
                _ds2 = slots_by_day.get(_d2.date, [])
                _uslots2 = [s2 for s2 in _ds2 if not _slot_is_exempt_daily(s2)]
                for _doc2 in doctors:
                    _dv2 = [_x2[(s2.slot_id, _doc2)] for s2 in _uslots2 if _x2.get((s2.slot_id, _doc2)) is not None]
                    if _dv2:
                        _model2.Add(sum(_dv2) <= 1)
            # fixed assignments (hard)
            for _fa2 in (fixed_assignments or []):
                try:
                    _fc2 = str(_fa2.get("column","")).strip().upper()
                    _fd2 = date.fromisoformat(str(_fa2.get("date","")).strip())
                    _fdc2 = norm_name(str(_fa2.get("doctor","")).strip())
                    if _fc2 == "J":
                        continue
                    _sf2 = next((s2 for s2 in slots_by_day.get(_fd2, []) if _fc2 in (s2.columns or [])), None)
                    if _sf2 and (_x2.get((_sf2.slot_id, _fdc2)) is not None):
                        _model2.Add(_x2[(_sf2.slot_id, _fdc2)] == 1)
                except Exception:
                    pass
            _model2.Minimize(sum(_extra2) if _extra2 else _model2.NewIntVar(0, 0, "z2"))
            _solver2 = cp_model.CpSolver()
            _solver2.parameters.max_time_in_seconds = 15.0
            _solver2.parameters.num_search_workers = 1
            _status2 = _solver2.Solve(_model2)
            if _status2 in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                # ── Retry granulare: trova quale gruppo causa l'infeasibility ──
                _groups: List[str] = []

                def _try_add(label: str, add_fn) -> bool:
                    """Ritorna True se aggiungere questo gruppo rende il modello infeasible."""
                    try:
                        import copy as _cp2
                        _m3 = cp_model.CpModel()
                        # Copia tutte le variabili e i vincoli del modello base
                        # (non possibile direttamente → testa il subset separatamente)
                        # Strategia: aggiungi il gruppo al modello2 già costruito e ri-solvi
                        _solver3 = cp_model.CpSolver()
                        _solver3.parameters.max_time_in_seconds = 8.0
                        _solver3.parameters.num_search_workers = 1
                        add_fn(_model2, _x2)
                        _st3 = _solver3.Solve(_model2)
                        if _st3 == cp_model.INFEASIBLE:
                            _groups.append(label)
                            return True
                        return False
                    except Exception as _e:
                        _groups.append(f"{label}(err:{_e})")
                        return True

                # Test 1: night spacing
                _night_pairs_added = []
                _ndvm2: Dict = {}
                for _s2 in slots:
                    if "J" in (_s2.columns or []):
                        for _dd2 in doctors:
                            if _x2.get((_s2.slot_id, _dd2)) is not None:
                                _ndvm2[(_s2.day.date, _dd2)] = _x2[(_s2.slot_id, _dd2)]
                _days_sorted2 = sorted(days, key=lambda d: d.date)
                _min_gap2 = int((cfg.get("global_constraints") or {}).get("night_spacing_days_min", 5) or 5)
                for _i2, _drow2 in enumerate(_days_sorted2):
                    for _k2 in range(1, _min_gap2):
                        _j2 = _i2 + _k2
                        if _j2 >= len(_days_sorted2):
                            break
                        _d1_2, _d2_2 = _drow2.date, _days_sorted2[_j2].date
                        for _doc2 in doctors:
                            _v1_2 = _ndvm2.get((_d1_2, _doc2))
                            _v2_2 = _ndvm2.get((_d2_2, _doc2))
                            if _v1_2 is not None and _v2_2 is not None:
                                _model2.Add(_v1_2 + _v2_2 <= 1)

                _solver_t1 = cp_model.CpSolver()
                _solver_t1.parameters.max_time_in_seconds = 8.0
                _st_t1 = _solver_t1.Solve(_model2)
                if _st_t1 == cp_model.INFEASIBLE:
                    _groups.append("night_spacing")

                # Test 2: university cap
                _gc3 = cfg.get("global_constraints") or {}
                _gc_uni3 = _gc3.get("university_doctors") or {}
                _uni_ratio3 = float(_gc3.get("university_ratio", 0.6))
                _working3 = sum(1 for d in days if d.dow in ["Mon","Tue","Wed","Thu","Fri","Sat"] and not is_festivo(d, cfg))
                _counts_as3 = cfg.get("pool_counts_as") or {}
                _uni_added = False
                for _udoc3, _udcfg3 in _gc_uni3.items():
                    _dn3 = norm_name(_udoc3)
                    if _dn3 not in doctors:
                        continue
                    _wterms3 = []
                    for _s3 in slots:
                        _v3 = _x2.get((_s3.slot_id, _dn3))
                        if _v3 is None:
                            continue
                        _w3 = max((_counts_as3.get(c, 1) for c in (_s3.columns or [])), default=1)
                        if _w3 > 0:
                            _wterms3.append(_w3 * _v3)
                    if not _wterms3:
                        continue
                    _tgt3 = round(_working3 * _uni_ratio3)
                    _uc3 = _model2.NewIntVar(0, len(_wterms3) * 2, f"uc3_{hash(_dn3)%10**6}")
                    _model2.Add(_uc3 == sum(_wterms3))
                    _model2.Add(_uc3 <= _tgt3 + 3)
                    _uni_added = True
                if _uni_added:
                    _slv3 = cp_model.CpSolver()
                    _slv3.parameters.max_time_in_seconds = 8.0
                    _st3 = _slv3.Solve(_model2)
                    if _st3 == cp_model.INFEASIBLE:
                        _groups.append("university_cap(target+3)")

                # Test 3: E/G hard cap
                _eg3_by_date: Dict = {}
                for _s3 in slots:
                    if _s3.rule_tag == "E_G":
                        _eg3_by_date.setdefault(_s3.day.date, _s3.slot_id)
                _egsl3 = [_x2.get((_sid, _dn)) for _sid in _eg3_by_date.values() for _dn in doctors if _x2.get((_sid, _dn)) is not None]
                # build per-doctor eg cap
                _rEG3 = (cfg.get("rules") or {}).get("E_G", {})
                _eg_pool3 = [norm_name(d) for d in (_rEG3.get("allowed") or []) if norm_name(d) in doctors]
                _n_eg3 = len(_eg3_by_date)
                if _eg_pool3 and _n_eg3 > 0:
                    import math as _math3
                    _n_act3 = max(sum(1 for d in _eg_pool3 if any(_x2.get((sid, d)) is not None for sid in _eg3_by_date.values())), 1)
                    _emax3 = _math3.ceil(_n_eg3 / _n_act3) + 1
                    for _d3 in _eg_pool3:
                        _ev3 = [_x2[(sid, _d3)] for sid in _eg3_by_date.values() if _x2.get((sid, _d3)) is not None]
                        if _ev3:
                            _ec3 = _model2.NewIntVar(0, _n_eg3, f"ec3_{hash(_d3)%10**6}")
                            _model2.Add(_ec3 == sum(_ev3))
                            _model2.Add(_ec3 <= _emax3)
                    _slv4 = cp_model.CpSolver()
                    _slv4.parameters.max_time_in_seconds = 8.0
                    _st4 = _slv4.Solve(_model2)
                    if _st4 == cp_model.INFEASIBLE:
                        _groups.append("EG_cap")

                # Test 4: night_off_next_day
                _gc4 = cfg.get("global_constraints") or {}
                _night_next4 = bool((_gc4.get("night_off") or {}).get("next_day", True))
                if _night_next4:
                    _days_s4 = sorted(days, key=lambda d: d.date)
                    for _i4, _drow4 in enumerate(_days_s4[:-1]):
                        _next4 = _days_s4[_i4 + 1].date
                        for _doc4 in doctors:
                            _v1_4 = _ndvm2.get((_drow4.date, _doc4))
                            if _v1_4 is None:
                                continue
                            for _s4 in slots_by_day.get(_next4, []):
                                if _slot_is_exempt_daily(_s4):
                                    continue
                                _v2_4 = _x2.get((_s4.slot_id, _doc4))
                                if _v2_4 is not None:
                                    _model2.Add(_v1_4 + _v2_4 <= 1)
                    _slv5 = cp_model.CpSolver()
                    _slv5.parameters.max_time_in_seconds = 10.0
                    _st5 = _slv5.Solve(_model2)
                    if _st5 == cp_model.INFEASIBLE:
                        _groups.append("night_off_next_day")

                # Test 5: H/I consecutive day
                _hi_ids5: Dict = {}
                for _s5 in slots:
                    for _c5 in (_s5.columns or []):
                        if _c5 in {"H", "I"}:
                            _hi_ids5.setdefault((_s5.day.date, _c5), []).append(_s5.slot_id)
                _ds5 = sorted(days, key=lambda d: d.date)
                for _i5 in range(len(_ds5) - 1):
                    _d1_5, _d2_5 = _ds5[_i5].date, _ds5[_i5 + 1].date
                    for _col5 in ("H", "I"):
                        for _doc5 in doctors:
                            if _doc5 == "Recupero":
                                continue
                            _vv1 = [_x2.get((sid, _doc5)) for sid in _hi_ids5.get((_d1_5, _col5), []) if _x2.get((sid, _doc5)) is not None]
                            _vv2 = [_x2.get((sid, _doc5)) for sid in _hi_ids5.get((_d2_5, _col5), []) if _x2.get((sid, _doc5)) is not None]
                            if _vv1 and _vv2:
                                _model2.Add(sum(_vv1) + sum(_vv2) <= 1)
                _slv6 = cp_model.CpSolver()
                _slv6.parameters.max_time_in_seconds = 10.0
                _st6 = _slv6.Solve(_model2)
                if _st6 == cp_model.INFEASIBLE:
                    _groups.append("HI_consecutive")

                if _groups:
                    _diag_hints.append(
                        f"Gruppo TROVATO [{'; '.join(_groups)}] causa infeasibility."
                    )
                else:
                    _diag_hints.append(
                        f"Tutti i gruppi OK — causa è in combinazione non testata "
                        f"(Recupero T Monday forced, L Recupero sum+short=3, o altro). "
                        f"Prossimo passo: disabilita pool_config e testa con YAML puro."
                    )
            elif _status2 == cp_model.INFEASIBLE:
                _diag_hints.append(
                    "INFEASIBLE anche senza quote/spacing: il problema è nella copertura base degli slot. "
                    "Controlla pool vuoti o fixed_assignments impossibili."
                )
            else:
                _diag_hints.append(f"Retry diagnostico: status={_status2} (timeout o unknown).")
        except Exception as _de:
            _diag_hints.append(f"Retry diagnostico fallito: {_de}")
        # Slot summary
        _slots_by_day_d: dict = {}
        for _s in slots:
            _slots_by_day_d.setdefault(_s.day.date.isoformat(), []).append(_s)
        _tight_days = []
        for _day_str, _day_slots in sorted(_slots_by_day_d.items()):
            _req = [_s for _s in _day_slots if _s.required]
            _empty = [_s for _s in _req if not _s.allowed]
            _very_small = [f"{'+'.join(_s.columns)}({len(_s.allowed)})" for _s in _req if 0 < len(_s.allowed) <= 2]
            if _empty or _very_small:
                _tight_days.append(f"{_day_str}: vuoti={['+'.join(_s.columns) for _s in _empty]} ristretti={_very_small}")
        _slots_diag = "; ".join(_tight_days) if _tight_days else "nessun slot vuoto"
        raise RuntimeError(
            f"{'; '.join(_diag_hints)} | Slot critici: {_slots_diag}"
        )
    # Identify required slots left blank (b_blank == 1) for diagnostics
    slots_by_id = {s.slot_id: s for s in slots}
    forced_blank_slots: List[str] = []
    forced_blank_details: List[Dict[str, object]] = []
    for sid, bv in blank_required_vars.items():
        try:
            if solver.Value(bv) == 1:
                forced_blank_slots.append(sid)
                s = slots_by_id.get(sid)
                if s is not None:
                    forced_blank_details.append({
                        "slot_id": sid,
                        "date": s.day.date.isoformat(),
                        "columns": list(s.columns),
                        "rule_tag": s.rule_tag,
                        "shift": s.shift,
                        "allowed_n": len(s.allowed or []),
                        "allowed": list(s.allowed or []),
                        "emergency_n": len(s.emergency_doctors or []),
                        "emergency_doctors": list(s.emergency_doctors or []),
                    })
        except Exception:
            pass

    assignment: Dict[str, Optional[str]] = {s.slot_id: None for s in slots}
    for s in slots:
        chosen = None
        for d in s.allowed:
            v = x.get((s.slot_id, d))
            if v is not None and solver.Value(v) == 1:
                chosen = d
                break
        assignment[s.slot_id] = chosen
    has_forced_blanks = bool(forced_blank_slots)
    stats = {
        "status": "OPTIMAL" if status == cp_model.OPTIMAL else "FEASIBLE",
        "objective": solver.ObjectiveValue(),
        "warnings": pre_solve_warnings,
        "lexicographic": {
            "coverage_status": _status_name(coverage_status),
            "coverage_cost": coverage_cost,
            "heavy_status": _status_name(heavy_status),
            "heavy_cost": heavy_cost,
            "final_status": _status_name(status),
        },
    }
    if has_forced_blanks:
        stats["status"] = "PARTIAL"
        stats["forced_blank_slots"] = sorted(forced_blank_slots)
        stats["forced_blank_details"] = sorted(
            forced_blank_details,
            key=lambda item: str(item.get("slot_id", "")),
        )
        stats.setdefault("warnings", []).append(
            f"{len(forced_blank_slots)} slot obbligatori lasciati vuoti (infeasible): "
            + ", ".join(sorted(forced_blank_slots))
        )

    # Diagnostics for the "2 lunedì" rule (do not affect feasibility).
    MonT_used_rec = 0
    if MonT_need_rec and MonT_target_dates:
        for d in MonT_target_dates:
            if assignment.get(f"{d}-T") == "Recupero":
                MonT_used_rec += 1
        stats["MonT_recupero"] = {
            "need": int(MonT_need_rec),
            "used": int(MonT_used_rec),
            "dates": [dd.isoformat() for dd in MonT_target_dates],
        }
        if MonT_used_rec != MonT_need_rec:
            stats.setdefault("warnings", []).append(
                f"T=Recupero su lunedì: attesi {MonT_need_rec}, ottenuti {MonT_used_rec}"
            )

    # Final layer: assign/reassign Reperibilità (C) using definitive rules
    assignment, cdiag = assign_reperibilita_C(cfg, days, slots, assignment)
    if isinstance(cdiag, dict):
        stats.update(cdiag)
    return assignment, stats
# -------------------------
# Fallback greedy (simple)
# -------------------------
def solve_greedy(cfg: dict, days: List[DayRow], slots: List[Slot]) -> Tuple[Dict[str, Optional[str]], Dict]:
    """
    Greedy with rule-aware priorities (fallback when OR-Tools is missing or after autorelax still infeasible).
    Goals:
    - always cover required slots when possible
    - enforce key HARD quotas (e.g., Recupero on L <= 3, Recupero on Y = 2 Mondays) and fixed-day rules
    - avoid obviously bad imbalances (especially nights)
    """
    # Rank shifts: Night first helps satisfy spacing/off constraints early
    shift_rank = {"Notte": 0, "Pomeriggio": 1, "Mattina": 2, "Any": 3}
    gc = cfg.get('global_constraints', {}) or {}
    daily_exempt_cols = {str(c).strip().upper() for c in (gc.get('daily_uniqueness_exempt_columns') or [])}
    def _slot_is_exempt_daily(s: Slot) -> bool:
        return any(str(c).strip().upper() in daily_exempt_cols for c in (s.columns or []))
    # Assign C (Reperibilità) as a final layer (so it never blocks core shifts)
    slots_non_c = [s for s in slots if s.columns != ['C']]
    
    # Sort: required first, then smallest candidate pool, then shift priority
    slots_sorted = sorted(
        slots_non_c,
        key=lambda s: (not s.required, len(s.allowed), shift_rank.get(s.shift, 9), s.rule_tag or "", s.slot_id),
    )
    assignment: Dict[str, Optional[str]] = {s.slot_id: None for s in slots}
    used_per_day: Dict[dt.date, Set[str]] = defaultdict(set)
    nights_by_doc: Dict[str, List[dt.date]] = defaultdict(list)
    load_total = Counter()
    load_by_tag: Dict[str, Counter] = defaultdict(Counter)
    gc = cfg.get("global_constraints", {}) or {}
    min_gap = int(gc.get("night_spacing_days_min", 5) or 5)
    night_off_next = bool((gc.get("night_off") or {}).get("next_day", True))
    # --- Hard quotas / caps from rules
    rules = cfg.get("rules", {}) or {}
    # J: weekend night hard max per doctor (e.g. Zito: 1)
    _rJ_greedy = rules.get("J") or {}
    wn_max_hard_greedy = {norm_name(k): int(v) for k, v in (_rJ_greedy.get("weekend_night_max_hard") or {}).items()}
    we_nights_by_doc: Dict[str, int] = defaultdict(int)
    # L: Recupero cap (interpret as MAX, not exact)
    L_cap_rec = None
    if isinstance(rules.get("L"), dict):
        L_cap_rec = rules["L"].get("quota_recupero_per_month")
        L_cap_rec = int(L_cap_rec) if L_cap_rec is not None else None
    L_rec_used = 0
    # Y: Monday specialist clinics
    # - Y_MAIN: always 1 doctor among the main pool (rotation)
    # - Y_REC: optional second line (Recupero) on exactly 2 Mondays/month
    Y_need_rec = 0
    if isinstance(rules.get("Y"), dict) and rules["Y"].get("recupero_two_mondays_per_month", False):
        Y_need_rec = 2
    Y_rec_used = 0
    Y_rec_slots_total = sum(1 for s in slots if getattr(s, "rule_tag", "") == "Y_REC")
    Y_rec_slots_done = 0
    Y_pool_counts = Counter()
    # Y affiancamento Recupero su T: i primi 2 lunedì del mese (fallback greedy)
    MonT_need_rec = 0
    MonT_target_dates: Set[dt.date] = set()
    if (
        isinstance(rules.get("Y"), dict)
        and rules["Y"].get("recupero_two_mondays_per_month", False)
        and rules["Y"].get("recupero_affianca_in_T", False)
    ):
        MonT_need_rec = 2
        monday_dates = sorted([d.date for d in days if d.dow == "Mon"])
        MonT_target_dates = set(monday_dates[:2])
    MonT_used_rec = 0

    # Night equalization target (if divisible)
    night_pool = set()
    if isinstance(rules.get("J"), dict):
        rJ = rules["J"]
        night_pool |= {norm_name(d) for d in (rJ.get("pool_other") or [])}
        night_pool |= {norm_name(d) for d in (rJ.get("monthly_quotas") or {}).keys()}
        night_pool.discard("Recupero")
    night_slots_total = sum(1 for s in slots if s.columns == ["J"])
    night_target = None
    if night_pool and night_slots_total > 0 and night_slots_total % len(night_pool) == 0:
        night_target = night_slots_total // len(night_pool)
    def can_assign(s: Slot, doc: str) -> bool:
        # per-day uniqueness (ignore placeholder 'Recupero')
        if (not _slot_is_exempt_daily(s)) and doc in used_per_day[s.day.date]:
            return False
        # L cap for Recupero
        if s.columns == ["L"] and doc == "Recupero" and L_cap_rec is not None and L_rec_used >= L_cap_rec:
            return False
        # Y_REC quota for Recupero (avoid exceeding)
        if (getattr(s, "rule_tag", "") == "Y_REC") and doc == "Recupero" and Y_need_rec and Y_rec_used >= Y_need_rec:
            return False
        # Night spacing
        if s.columns == ["J"]:
            for prev in nights_by_doc[doc]:
                if abs((s.day.date - prev).days) < min_gap:
                    return False
            # Weekend night hard max (e.g. Zito: max 1 weekend night)
            if s.day.dow in ("Sat", "Sun") and doc in wn_max_hard_greedy:
                if we_nights_by_doc[doc] >= wn_max_hard_greedy[doc]:
                    return False
        # Night off next day (if doc did night previous day)
        if night_off_next:
            prev_day = s.day.date - dt.timedelta(days=1)
            if prev_day in nights_by_doc[doc]:
                return False
        # K no consecutive days (if enabled in rules)
        if "K" in s.columns and isinstance(rules.get("K"), dict) and rules["K"].get("no_consecutive_days_same_doctor", False):
            prev_day = s.day.date - dt.timedelta(days=1)
            if assignment.get(f"{prev_day}-K") == doc:
                return False
        # D/F different is handled by per-day uniqueness + separate slots; ok.
        return True
    def score_candidate(s: Slot, doc: str) -> Tuple:
        """
        Lower is better.
        Prioritize:
        - Night balance (nights first, then total load)
        - Column-specific balance (e.g., Y rotation)
        - Total load balance
        """
        tag = s.rule_tag or ""
        if s.columns == ["J"]:
            night_cnt = len(nights_by_doc[doc])
            # keep closer to target if known
            tgt_pen = abs(night_cnt - (night_target or 0)) if night_target is not None else night_cnt
            return (tgt_pen, night_cnt, load_total[doc], doc.lower())
        if (s.rule_tag or "") == "Y_REC":
            return (0, load_total[doc], doc.lower())
        if (s.rule_tag or "") == "Y_MAIN":
            return (Y_pool_counts[doc], load_total[doc], doc.lower())
        # default: balance by tag then total
        return (load_by_tag[tag][doc], load_total[doc], doc.lower())
    def pick(s: Slot) -> Optional[str]:
        candidates = [d for d in s.allowed if can_assign(s, d)]
        if not candidates:
            return None
        # Forza/nega Recupero su T (lunedi) in base ai 2 lunedi target (fallback greedy)
        nonlocal MonT_used_rec
        if s.columns == ["T"] and getattr(s.day, "dow", "") == "Mon" and MonT_need_rec:
            if s.day.date in MonT_target_dates:
                if "Recupero" in candidates:
                    return "Recupero"
            else:
                # fuori dai 2 lunedi target: evita Recupero se ci sono alternative
                if "Recupero" in candidates and len(candidates) > 1:
                    candidates = [d for d in candidates if d != "Recupero"]

        # Force Recupero on Y_REC when needed and we're running out of Monday slots
        nonlocal Y_rec_slots_done, Y_rec_used
        if (getattr(s, "rule_tag", "") == "Y_REC") and Y_need_rec:
            remaining_after_this = (Y_rec_slots_total - (Y_rec_slots_done + 1))
            remaining_need = (Y_need_rec - Y_rec_used)
            if remaining_need <= 0:
                return None  # leave blank
            if remaining_need > remaining_after_this:
                # must use Recupero now (if possible)
                if "Recupero" in candidates:
                    return "Recupero"
                return None
            # otherwise, leave blank to keep flexibility
            return None
        candidates.sort(key=lambda d: score_candidate(s, d))
        return candidates[0]
    conflicts = []
    for s in slots_sorted:
        if (s.rule_tag or "") == "Y_REC":
            Y_rec_slots_done += 1
        chosen = pick(s)
        if chosen is None:
            if s.required:
                conflicts.append(f"UNFILLED required slot {s.slot_id} ({s.columns})")
            continue
        assignment[s.slot_id] = chosen
        if not _slot_is_exempt_daily(s):
            used_per_day[s.day.date].add(chosen)
        load_total[chosen] += 1
        load_by_tag[s.rule_tag or ""][chosen] += 1
        if s.columns == ["J"]:
            nights_by_doc[chosen].append(s.day.date)
            if s.day.dow in ("Sat", "Sun"):
                we_nights_by_doc[chosen] += 1
        if s.columns == ["L"] and chosen == "Recupero":
            L_rec_used += 1
        if s.columns == ["T"] and getattr(s.day, "dow", "") == "Mon" and chosen == "Recupero":
            MonT_used_rec += 1
        if (s.rule_tag or "") == "Y_REC":
            if chosen == "Recupero":
                Y_rec_used += 1
        if (s.rule_tag or "") == "Y_MAIN":
            Y_pool_counts[chosen] += 1
    # Final layer: assign/reassign Reperibilità (C) using definitive rules
    assignment, cdiag = assign_reperibilita_C(cfg, days, slots, assignment)
    # restore C slots as required, even if greedy did not assign them earlier
    stats = {
        "status": "GREEDY",
        "conflicts": conflicts,
        "loads": dict(load_total),
        "nights_per_doc": {k: len(v) for k, v in nights_by_doc.items()},
        "L_recupero_used": L_rec_used,
        "Y_recupero_used": Y_rec_used,
        "T_recupero_mondays_used": MonT_used_rec,
    }
    if isinstance(cdiag, dict):
        stats.update(cdiag)
    return assignment, stats
# -------------------------
# Write output Excel
# -------------------------
def write_output(
    wb: openpyxl.Workbook,
    ws: openpyxl.worksheet.worksheet.Worksheet,
    days: List[DayRow],
    slots: List[Slot],
    assignment: Dict[str, Optional[str]],
    out_path: Path,
    cfg: Optional[dict] = None,
    unav_map: Optional[Dict[str, Dict[dt.date, Set[str]]]] = None,
):
    disabled_cols = {
        str(c).strip().upper()
        for c in ((cfg or {}).get("disabled_columns") or [])
        if str(c).strip()
    }
    # Clear target columns (only those managed)
    managed_cols=set()
    for s in slots:
        managed_cols |= set(s.columns)
    managed_cols |= disabled_cols
    # do not wipe A,B headers; wipe from row 2
    if cfg and isinstance(cfg.get("rules", {}), dict) and "AA" in (cfg.get("rules") or {}) and "AA" not in disabled_cols:
        managed_cols.add("AA")
    for drow in days:
        for col in managed_cols:
            ws[f"{col}{drow.row_idx}"].value = None
    # Fill (support multiple assignments on the same cell, e.g. Y_MAIN + Y_REC)
    slot_by_id = {s.slot_id: s for s in slots}
    assigned_by_day: Dict[dt.date, Set[str]] = defaultdict(set)
    cell_values: Dict[Tuple[int, str], List[str]] = defaultdict(list)
    # Iterate slots in creation order for stable writing
    for s in slots:
        doc = assignment.get(s.slot_id)
        if not doc:
            continue
        for col in s.columns:
            cell_values[(s.day.row_idx, col)].append(doc)
        assigned_by_day[s.day.date].add(doc)
    
    # Post-process: affiancamento Recupero in Y sui 2 lunedì fissi (i primi 2 del mese) – vedi vincolo su T.
    if cfg and isinstance(cfg.get("rules", {}), dict) and "Y" not in disabled_cols:
        rY = (cfg.get("rules") or {}).get("Y") or {}
        if rY.get("recupero_two_mondays_per_month", False) and rY.get("recupero_affianca_in_T", False):
            monday_dates = [drow.date for drow in days if getattr(drow, "dow", "") == "Mon"]
            monday_dates = sorted(monday_dates)[:2]
            target = set(monday_dates)
            for drow in days:
                if drow.date in target and assignment.get(f"{drow.date}-T") == "Recupero":
                    cell_values[(drow.row_idx, "Y")].append("Recupero")

    for (row_idx, col), docs in cell_values.items():
        # de-duplicate while preserving order
        seen = set()
        uniq = [d for d in docs if not (d in seen or seen.add(d))]
        # Y può legittimamente avere Recupero+medico (affiancamento): usa \n
        # V il venerdì ha legittimamente 2 medici (Crea + Dattilo/Allegra): usa \n
        # Tutte le altre colonne: deve esserci un solo medico; se ce ne sono due è un bug → primo
        if col in ("Y", "V"):
            ws[f"{col}{row_idx}"].value = "\n".join(uniq)
        else:
            ws[f"{col}{row_idx}"].value = uniq[0] if uniq else None

    # Ensure headers exist for new/optional columns (Z, AA) even on older templates.
    for _col, _label in [
        ("Z", ((cfg or {}).get("columns") or {}).get("Z", "Vascolare")),
        ("AA", ((cfg or {}).get("columns") or {}).get("AA", "SPOC")),
    ]:
        try:
            h = ws[f"{_col}1"]
            if h.value is None or str(h.value).strip() == "":
                h.value = _label
        except Exception:
            pass

    # Backward-compat: older templates may only have AD/AE (or no headers at all).
    # Ensure headers exist for the "medici liberi" block.
    for _col, _label in [
        ("AD", "Medici liberi 1"),
        ("AE", "Medici liberi 2"),
        ("AF", "Medici liberi 3"),
        ("AG", "Medici liberi 4"),
    ]:
        try:
            h = ws[f"{_col}1"]
            if h.value is None or str(h.value).strip() == "":
                h.value = _label
        except Exception:
            pass
    # Fill medici liberi 1/2/3/4 (AD/AE/AF/AG)
    # Nota: le colonne AD/AE possono rimanere vuote. Se però inseriamo un nome,
    # deve essere un medico *disponibile* quel giorno (nessuna indisponibilità registrata).
    # Base roster: unione dei medici noti da YAML (pools/fixed/unavailability),
    # NON dipende dal dominio dei singoli slot (che può diventare vuoto per indisponibilità).
    all_docs = set(collect_doctors(cfg))
    # 'Recupero' è un medico a tutti gli effetti: può comparire anche tra i liberi.
    unav_map = unav_map or {}
    # Escludi d'ufficio lo SMONTANTE NOTTE: chi ha fatto la NOTTE (colonna J) il giorno prima
    # non può essere considerato "libero" il giorno successivo (anche se non assegnato a nessuna colonna).
    night_by_date: Dict[dt.date, str] = {}
    for s in slots:
        if s.columns == ["J"]:
            nd = assignment.get(s.slot_id)
            if nd:
                night_by_date[s.day.date] = nd

    # Pre-calcola, per ogni giorno, quali fasce (M/P/N) ogni medico può effettivamente
    # coprire in base ai pool degli slot — esclude chi non è mai in un pool di notte, ecc.
    _SHIFT_TO_KEY = {"Mattina": "M", "Pomeriggio": "P", "Notte": "N",
                     "Diurno": None}  # Diurno → M e P
    doc_eligible_by_date: Dict[dt.date, Dict[str, Set[str]]] = {}
    for s in slots:
        s_date = s.day.date
        s_shift = getattr(s, "shift", "") or ""
        if s_shift == "Any":
            continue  # C reperibilità — ignora per il label
        keys: List[str] = []
        if s_shift == "Diurno":
            keys = ["M", "P"]
        elif s_shift in _SHIFT_TO_KEY and _SHIFT_TO_KEY[s_shift]:
            keys = [_SHIFT_TO_KEY[s_shift]]
        else:
            continue
        if s_date not in doc_eligible_by_date:
            doc_eligible_by_date[s_date] = {}
        for doc in (s.allowed or []):
            if doc not in doc_eligible_by_date[s_date]:
                doc_eligible_by_date[s_date][doc] = set()
            doc_eligible_by_date[s_date][doc].update(keys)

    # Per i giorni festivi con assegnazione forzata (sorteggio), s.allowed viene ridotto
    # a [solo_quel_medico], quindi gli altri medici del pool festivi spariscono da
    # doc_eligible_by_date → non appaiono in Medici liberi. Aggiungiamo il pool completo.
    if cfg and isinstance(cfg.get("rules"), dict):
        rFest = (cfg["rules"].get("Festivi") or {})
        _fp_m = [norm_name(d) for d in (rFest.get("pool_mattina") or rFest.get("pool") or []) if d]
        _fp_p = [norm_name(d) for d in (rFest.get("pool_pomeriggio") or rFest.get("pool") or []) if d]
        for drow in days:
            if not is_festivo(drow, cfg):
                continue
            fdate = drow.date
            if fdate not in doc_eligible_by_date:
                doc_eligible_by_date[fdate] = {}
            for doc in _fp_m:
                doc_eligible_by_date[fdate].setdefault(doc, set()).add("M")
            for doc in _fp_p:
                doc_eligible_by_date[fdate].setdefault(doc, set()).add("P")

    # Medici in never_in_J non fanno mai notti → togli "N" dal loro eligible
    # (l'emergency expansion di J li aggiunge come fallback ma non possono davvero fare J)
    if cfg and isinstance(cfg.get("rules"), dict):
        _j_never = {norm_name(d) for d in ((cfg["rules"].get("J") or {}).get("never_in_J") or ["De Gregorio", "Manganaro"])}
        for _elig_map in doc_eligible_by_date.values():
            for _dn in _j_never:
                if _dn in _elig_map:
                    _elig_map[_dn].discard("N")

    def _free_label(doc: str, unav_shifts: Set[str], eligible: Set[str]) -> Optional[str]:
        """Restituisce la stringa da scrivere in Medici liberi, o None se non disponibile.
        - eligible: fasce (M/P/N) per cui il medico è in almeno un pool in quel giorno
        - Nessuna indisponibilità e tutto eligible → solo il nome
        - Indisponibilità totale (Any / Tutto il giorno) → None
        - Indisponibilità parziale o pool parziale → "Nome (fasce_libere)"
        """
        if "Any" in unav_shifts or "Tutto il giorno" in unav_shifts:
            return None
        # Espandi fasce indisponibili
        unav_exp: Set[str] = set()
        for sh in unav_shifts:
            if sh == "Mattina":
                unav_exp.add("M")
            elif sh == "Pomeriggio":
                unav_exp.add("P")
            elif sh == "Notte":
                unav_exp.add("N")
            elif sh == "Diurno":
                unav_exp.update({"M", "P"})
        # Fasce effettivamente disponibili = eligible - indisponibili
        avail = [k for k in ["M", "P", "N"] if k in eligible and k not in unav_exp]
        if not avail:
            return None
        if avail == [k for k in ["M", "P", "N"] if k in eligible]:
            # Tutte le fasce eligibili sono disponibili
            if not unav_shifts:
                return doc  # nessuna indisponibilità → no suffisso
            # Ha indisponibilità ma non nelle fasce del suo pool → mostra comunque
            if eligible == {"M", "P", "N"}:
                return doc
        abbr_str = ", ".join(avail)
        # Mostra suffisso solo se non ha tutte e 3 le fasce o ha indisponibilità
        all_three = {"M", "P", "N"}
        if eligible >= all_three and not unav_shifts:
            return doc
        if set(avail) == eligible and not unav_shifts:
            return doc
        return f"{doc} ({abbr_str})"

    for drow in days:
        assigned_today = assigned_by_day.get(drow.date, set())
        smontante = night_by_date.get(drow.date - dt.timedelta(days=1))
        smontanti = {smontante} if smontante else set()
        _eligible_today = doc_eligible_by_date.get(drow.date, {})
        free_full: List[str] = []   # completamente disponibili
        free_partial: List[str] = []  # parzialmente disponibili (con suffisso)
        for doc in sorted(all_docs, key=lambda s: s.lower()):
            if doc in assigned_today or doc in smontanti:
                continue
            unav_shifts = unav_map.get(doc, {}).get(drow.date, set())
            eligible = _eligible_today.get(doc, set())
            if not eligible:
                continue  # medico senza pool attivo quel giorno → non compare
            label = _free_label(doc, unav_shifts, eligible)
            if label is None:
                continue
            if "(" in label:
                free_partial.append(label)
            else:
                free_full.append(label)
        free = free_full + free_partial
        ws[f"AD{drow.row_idx}"].value = free[0] if len(free) > 0 else None
        ws[f"AE{drow.row_idx}"].value = free[1] if len(free) > 1 else None
        ws[f"AF{drow.row_idx}"].value = free[2] if len(free) > 2 else None
        ws[f"AG{drow.row_idx}"].value = free[3] if len(free) > 3 else None

    # Fill AA (SPOC): solo Lun/Mer, copiando il medico di K o T. Il bilanciamento è fatto
    # in post-process scegliendo (quando K != T) il candidato meno usato nel mese.
    if cfg and "rules" in cfg and "AA" in (cfg.get("rules") or {}) and "AA" not in disabled_cols:
        rAA = (cfg.get("rules") or {}).get("AA") or {}
        copy_from = [str(c).strip().upper() for c in (rAA.get("copy_from") or ["K", "T"])]
        counts_by_month: Dict[Tuple[int, int], Dict[str, int]] = defaultdict(lambda: defaultdict(int))
        for drow in days:
            if not dayspec_contains(drow.dow, rAA.get("days")):
                continue
            candidates: List[str] = []
            for src in copy_from:
                try:
                    v = ws[f"{src}{drow.row_idx}"].value
                except Exception:
                    v = None
                if v is None:
                    continue
                name = str(v).splitlines()[0].strip()
                if name:
                    candidates.append(norm_name(name))
            # uniq preserve order
            seen=set()
            candidates=[c for c in candidates if c and not (c in seen or seen.add(c))]
            candidates=[c for c in candidates if c != "Recupero"]
            if not candidates:
                continue
            if len(candidates) == 1:
                chosen = candidates[0]
            else:
                mkey = (drow.date.year, drow.date.month)
                chosen = min(candidates, key=lambda c: (counts_by_month[mkey].get(c, 0), c.lower()))
            ws[f"AA{drow.row_idx}"].value = chosen
            mkey = (drow.date.year, drow.date.month)
            counts_by_month[mkey][chosen] += 1

    # AB: the template may pre-color AB cells on Saturdays. Since AB is only required
    # on Thursdays and on exactly N Saturdays (filled via the optional AB_SAT slot),
    # clear any pre-existing fill on Saturdays when AB is intentionally blank.
    try:
        no_fill = PatternFill()
        for drow in days:
            if getattr(drow, "dow", "") != "Sat":
                continue
            cell = ws[f"AB{drow.row_idx}"]
            v = cell.value
            if v is None or (isinstance(v, str) and v.strip() == ""):
                cell.fill = no_fill
    except Exception:
        pass

    # Highlight blanks in yellow:
    # - optional slots with blank_penalty > 0 (relief valves, empty domains)
    # - ALL blank slots for non-indispensable columns (se critical_services configurato)
    yellow_fill = PatternFill(fill_type="solid", start_color="FFFFF2CC", end_color="FFFFF2CC")
    _critical_cols = set((cfg.get("pool_critical_services") or {}).keys()) if cfg else set()
    _indispensable_mode = bool(_critical_cols)  # attivo solo se l'utente ha configurato critical_services
    for s in slots:
        if assignment.get(s.slot_id) is not None:
            continue
        bp = int(getattr(s, "blank_penalty", 0) or 0)
        _non_indispensable = _indispensable_mode and not any(c in _critical_cols for c in (s.columns or []))
        if bp <= 0 and not _non_indispensable:
            continue
        for col in (s.columns or []):
            cell = ws[f"{col}{s.day.row_idx}"]
            v = cell.value
            if v is None or (isinstance(v, str) and v.strip() == ""):
                cell.fill = yellow_fill
    _write_riepilogo_sheet(wb, ws, days, slots, assignment, cfg)
    wb.save(out_path)

def _write_riepilogo_sheet(
    wb: openpyxl.Workbook,
    ws_main: openpyxl.worksheet.worksheet.Worksheet,
    days: List[DayRow],
    slots: List[Slot],
    assignment: Dict[str, Optional[str]],
    cfg: Optional[dict] = None,
) -> None:
    """Aggiunge/aggiorna il foglio 'Riepilogo' con conteggi turni per medico.

    I conteggi sono formule Excel (COUNTIFS) che si aggiornano automaticamente
    quando l'admin modifica nomi nel foglio principale e salva.

    Colonna helper nascosta: contiene flag festivo (1) / feriale (0) per ogni
    riga-giorno — statica (calcolata dalle date), non cambia con i nomi.

    Peso per giorno: J (notte) = 2 | C (reperibilità) non conta | altro = 1.
    Obiettivo universitari usa lo stesso peso ma con J=1 se night_counts_double=False.
    """
    cfg = cfg or {}
    col_map: Dict[str, str] = cfg.get("columns") or {}
    SKIP_COLS = {"AD", "AE", "AF", "AG"}
    op_cols = [c for c in col_map if c not in SKIP_COLS]

    # ── Festivi ────────────────────────────────────────────────────────────
    year = days[0].date.year if days else dt.date.today().year
    holidays = italy_public_holidays(year)
    extra_hol: Set[dt.date] = set()
    for _x in (cfg.get("festivi_extra") or []):
        try:
            extra_hol.add(parse_date(_x))
        except Exception:
            pass
    day_info = {d.date: d for d in days}

    def _is_fes(date_: dt.date) -> bool:
        d = day_info.get(date_)
        return d is not None and (d.dow == "Sun" or date_ in holidays or date_ in extra_hol)

    # ── University doctors config ──────────────────────────────────────────
    gc = cfg.get("global_constraints") or {}
    gc_uni = gc.get("university_doctors") or {}
    uni_ratio = float(gc.get("university_ratio", 0.6))
    pct = int(uni_ratio * 100)
    working_days_n = sum(1 for d in days if d.dow in ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat") and not is_festivo(d, cfg))
    uni_target = round(working_days_n * uni_ratio)
    uni_docs = {norm_name(k) for k in gc_uni}
    uni_night_double = {
        norm_name(k): bool((v or {}).get("night_counts_double", False))
        for k, v in gc_uni.items()
    }

    # ── All assigned doctors (list determines Riepilogo rows) ─────────────
    all_docs: Set[str] = set()
    for s in slots:
        doc = assignment.get(s.slot_id)
        if doc:
            all_docs.add(doc)

    # ── Styles ────────────────────────────────────────────────────────────
    bold = openpyxl.styles.Font(bold=True)
    title_font = openpyxl.styles.Font(bold=True, size=13)
    hdr_fill = PatternFill(fill_type="solid", start_color="FFD9E1F2", end_color="FFD9E1F2")
    uni_fill = PatternFill(fill_type="solid", start_color="FFFFE2CC", end_color="FFFFE2CC")
    center = openpyxl.styles.Alignment(horizontal="center", wrap_text=True)

    mese_label = days[0].date.strftime("%B %Y") if days else ""
    n_hdr_cols = len(op_cols) + 5  # Medico + op_cols + Feriali + Festivi + Totale + Obiettivo

    # ── Create Riepilogo as second sheet (right after main sheet) ─────────
    sname = "Riepilogo"
    if sname in wb.sheetnames:
        del wb[sname]
    main_idx = wb.sheetnames.index(ws_main.title)
    ws = wb.create_sheet(sname, main_idx + 1)

    # ── Formula building blocks ────────────────────────────────────────────
    first_row = days[0].row_idx   # main sheet row of first day (= 2)
    last_row = days[-1].row_idx

    # Escape sheet name for cross-sheet formula reference
    ms_name = ws_main.title
    ms_ref = f"'{ms_name}'" if any(c in ms_name for c in (" ", "'", "!", "[", "]")) else ms_name

    # Hidden helper column: festivo flags at the same row indices as the main sheet.
    # Riepilogo!$AH$2:$AH$32  ←→  main sheet rows 2..32 (one per day).
    helper_col_idx = n_hdr_cols + 3
    hlp_letter = get_column_letter(helper_col_idx)
    hlp_range = f"Riepilogo!${hlp_letter}${first_row}:${hlp_letter}${last_row}"

    def _main_range(col: str) -> str:
        return f"{ms_ref}!${col}${first_row}:${col}${last_row}"

    def _cifs(col: str, dc: str, flag: int) -> str:
        """COUNTIFS: doctor in col filtered by feriale(0)/festivo(1).
        Wildcard (*name*) handles multi-doctor cells (e.g. V on Friday)."""
        return f'COUNTIFS({_main_range(col)},"*"&{dc}&"*",{hlp_range},{flag})'

    def _cif(col: str, dc: str) -> str:
        """COUNTIF (no flag): total occurrences of doctor in col."""
        return f'COUNTIF({_main_range(col)},"*"&{dc}&"*")'

    def col_display_formula(col: str, r: int) -> str:
        """Returns cell formula producing 'N+Mf' | 'N' | 'Mf' | ''."""
        dc = f"$A{r}"
        fer = _cifs(col, dc, 0)
        fes = _cifs(col, dc, 1)
        return (
            f'=IF(AND({fer}=0,{fes}=0),"",'
            f'IF({fes}=0,{fer},'
            f'IF({fer}=0,{fes}&"f",{fer}&"+"&{fes}&"f")))'
        )

    def weighted_formula(r: int, j_coeff: int, fes_flag: Optional[int]) -> str:
        """Numeric weighted-turni formula. C always excluded.
        Usa SUMPRODUCT+ISNUMBER(SEARCH()) per riga per evitare il doppio conteggio
        dei medici che coprono più colonne nello stesso giorno (D+F share, slot EG,
        DE/HI festivi, K+AA copia). Ogni giornata lavorativa conta 1, J conta j_coeff.
        j_coeff: coefficient for J (2 = standard, 1 = university without night_double).
        fes_flag: 0=feriali only, 1=festivi only, None=all."""
        dc = f"$A{r}"
        non_j_cols = [c for c in op_cols if c != "C" and c != "J"]
        has_j = "J" in op_cols

        # Filtro festivo/feriale (la colonna helper contiene 0=feriale, 1=festivo)
        if fes_flag is None:
            filt = ""
        else:
            filt = f"*({hlp_range}={fes_flag})"

        parts = []

        # Colonne non-J: presenza per riga → 1 per giornata lavorativa
        # (deduplica slot multi-colonna: D+F share, EG, DE festivo, HI festivo, K+AA, ecc.)
        if non_j_cols:
            search_terms = "+".join(
                f"ISNUMBER(SEARCH({dc},{_main_range(c)}))"
                for c in non_j_cols
            )
            parts.append(f"SUMPRODUCT((({search_terms})>0)*1{filt})")

        # Colonna J: presenza per riga × j_coeff
        if has_j:
            parts.append(f"SUMPRODUCT(ISNUMBER(SEARCH({dc},{_main_range('J')}))*{j_coeff}{filt})")

        return ("=" + "+".join(parts)) if parts else "=0"

    # ── Row 1: title ───────────────────────────────────────────────────────
    write_row = 1
    ws.cell(write_row, 1, f"Riepilogo Turni – {mese_label}").font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_hdr_cols)

    # ── Row 2: headers ─────────────────────────────────────────────────────
    write_row = 2
    hdr = (
        ["Medico"]
        + [col_map.get(c, c) for c in op_cols]
        + ["Feriali (peso)", "Festivi (peso)", "Totale (peso)", "Obiettivo *"]
    )
    for ci, val in enumerate(hdr, 1):
        cell = ws.cell(write_row, ci, val)
        cell.font = bold
        cell.fill = hdr_fill
        cell.alignment = center

    # Column indices for the summary columns
    fer_col_idx = len(op_cols) + 2
    fes_col_idx = fer_col_idx + 1
    tot_col_idx = fes_col_idx + 1
    obj_col_idx = tot_col_idx + 1
    fer_letter = get_column_letter(fer_col_idx)
    fes_letter = get_column_letter(fes_col_idx)

    # ── Rows 3+: one per doctor ────────────────────────────────────────────
    for doc in sorted(all_docs):
        write_row += 1
        r = write_row
        doc_n = norm_name(doc)
        is_uni = doc_n in uni_docs
        night_double = uni_night_double.get(doc_n, False)

        # A: doctor name — static, used by all formulas in this row via $A{r}
        ws.cell(r, 1, doc)

        # Per-column display: text formula "N+Mf" (updates live)
        for ci, col in enumerate(op_cols, 2):
            cell = ws.cell(r, ci)
            cell.value = col_display_formula(col, r)
            cell.alignment = center

        # Feriali (peso), Festivi (peso): numeric COUNTIFS formulas
        ws.cell(r, fer_col_idx).value = weighted_formula(r, j_coeff=2, fes_flag=0)
        ws.cell(r, fes_col_idx).value = weighted_formula(r, j_coeff=2, fes_flag=1)
        # Totale = Feriali + Festivi
        ws.cell(r, tot_col_idx).value = f"={fer_letter}{r}+{fes_letter}{r}"

        # Obiettivo: only for university doctors
        if is_uni:
            j_uni = 2 if night_double else 1
            uni_body = weighted_formula(r, j_coeff=j_uni, fes_flag=None)[1:]  # strip leading "="
            ws.cell(r, obj_col_idx).value = f'=TEXT({uni_body},"0")&"/{uni_target} ({pct}%)"'
            for ci in range(1, n_hdr_cols + 1):
                ws.cell(r, ci).fill = uni_fill

    # ── Notes ─────────────────────────────────────────────────────────────
    write_row += 2
    ws.cell(write_row, 1, "Note:").font = bold
    notes = [
        "  • Peso per giorno: J (notte) = 2 | C (reperibilità) non conta | altro = 1 per giornata lavorativa",
        "  • Turni multipli nella stessa giornata (es. D+F share, slot EG, festivi DE/HI, K+AA) contano 1",
        "  • Formato celle colonne: N = feriali  |  N+Mf = N feriali + M festivi  |  Mf = solo festivo",
        "  • Festivi = domeniche + festivi nazionali italiani",
        f"  • Giorni lavorativi lun-sab del mese: {working_days_n}",
        f"  • Obiettivo universitari (arancione): /{uni_target} = {working_days_n}×{pct}%"
        f"  (J={'2' if any(uni_night_double.values()) else '1'} se night_counts_double, C esclusa)",
        f"  • Universitari: {', '.join(sorted(uni_docs))}",
        "  ⚠ I conteggi si aggiornano automaticamente modificando i nomi nel foglio principale.",
    ]
    for note in notes:
        write_row += 1
        ws.cell(write_row, 1, note)

    # ── Helper column: festivo flags ───────────────────────────────────────
    # Written last to avoid interfering with write_row tracking.
    # Row alignment: day.row_idx == row index in this helper column
    # (both start at 2), so COUNTIFS ranges align correctly.
    for day in days:
        ws.cell(day.row_idx, helper_col_idx, 1 if _is_fes(day.date) else 0)
    ws.column_dimensions[hlp_letter].hidden = True

    # ── Column widths ──────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 18
    for i in range(2, n_hdr_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13


def write_solver_log(out_path: Path, stats: Dict) -> Optional[Path]:
    """
    Writes a human-readable solver log next to the output Excel.
    Includes: per-month status, objective, relief valves used, and day-level bottlenecks if any.
    """
    try:
        log_path = out_path.with_name(out_path.stem + "_solverlog.txt")
        lines: List[str] = []
        lines.append(f"Output: {out_path}")
        lines.append(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}")
        lines.append(f"Overall status: {stats.get('status')}")
        second_look = stats.get("heavy_festive_second_look") or []
        if second_look:
            lines.append("")
            lines.append("Second-look pesanti festivi/weekend applicato:")
            for item in second_look:
                cols = "+".join(item.get("columns") or [])
                lines.append(
                    f"  - {item.get('date')} {cols}: "
                    f"{item.get('from')} -> {item.get('to')}"
                )
        months = stats.get("months") or {}
        for mk in sorted(months.keys()):
            sm = months.get(mk) or {}
            lines.append("")
            lines.append(f"== {mk} ==")
            lines.append(f"status: {sm.get('status')}")
            if "objective" in sm:
                lines.append(f"objective: {sm.get('objective')}")
            lex = sm.get("lexicographic") or {}
            if lex:
                lines.append(
                    "lexicographic: "
                    f"coverage={lex.get('coverage_status')} cost={lex.get('coverage_cost')}; "
                    f"heavy={lex.get('heavy_status')} cost={lex.get('heavy_cost')}; "
                    f"final={lex.get('final_status')}"
                )
            if sm.get("autorelax"):
                lines.append(f"autorelax: {sm.get('autorelax')}")
            if sm.get("solver_error"):
                lines.append(f"solver_error: {sm.get('solver_error')}")
            j_overrides = sm.get("j_blank_week_overrides") or {}
            if j_overrides:
                lines.append("Eccezioni J applicate dalla GUI:")
                for wk_key in sorted(j_overrides.keys()):
                    vals = j_overrides.get(wk_key) or []
                    lines.append(f"  - {wk_key}: {', '.join(vals) if vals else 'nessuna J vuota'}")
            # Slot obbligatori lasciati bianchi (PARTIAL)
            fbs = sm.get("forced_blank_slots") or []
            if fbs:
                lines.append(f"ATTENZIONE: {len(fbs)} slot obbligatori NON coperti (infeasible parziale):")
                details_by_slot = {
                    str(item.get("slot_id")): item
                    for item in (sm.get("forced_blank_details") or [])
                    if item.get("slot_id")
                }
                for sid in fbs:
                    detail = details_by_slot.get(str(sid)) or {}
                    if detail:
                        allowed = detail.get("allowed") or []
                        emergency = detail.get("emergency_doctors") or []
                        allowed_preview = ", ".join(map(str, allowed[:12]))
                        emergency_preview = ", ".join(map(str, emergency[:8]))
                        suffix = (
                            f" | tag={detail.get('rule_tag') or '-'}"
                            f" | allowed_n={detail.get('allowed_n')}"
                        )
                        if allowed_preview:
                            suffix += f" | allowed={allowed_preview}"
                            if len(allowed) > 12:
                                suffix += ", ..."
                        if emergency_preview:
                            suffix += f" | emergenza={emergency_preview}"
                            if len(emergency) > 8:
                                suffix += ", ..."
                        lines.append(f"  - {sid}{suffix}")
                    else:
                        lines.append(f"  - {sid}")
            # Relief used
            ru = sm.get("relief_used") or {}
            if ru.get("kt_share_days") or ru.get("blank_columns"):
                if ru.get("kt_share_days"):
                    lines.append("relief: K=T same doctor days: " + ", ".join(ru.get("kt_share_days")))
                bc = ru.get("blank_columns") or {}
                for c in sorted(bc.keys()):
                    lines.append(f"relief: blank {c} on: " + ", ".join(bc[c]))
            # Day-level bottlenecks (if OR-Tools failed at least once)
            bl = sm.get("day_level_bottlenecks") or []
            if bl:
                lines.append("")
                lines.append("Day-level bottlenecks (ignores cross-day constraints):")
                for item in bl[:10]:
                    lines.append(f"- {item.get('date')} ({item.get('dow')}): required_slots={item.get('required_slots')}, union_doctors={item.get('union_doctors')}")
                    for us in (item.get("unmatched_slots") or [])[:3]:
                        lines.append(f"    * {us.get('slot_id')} cols={us.get('columns')} allowed_n={us.get('allowed_n')}")
            # Day-by-day diagnostic
            dd = sm.get("daily_diagnostic") or []
            if dd:
                lines.append("")
                lines.append("== Diagnostica giorno per giorno ==")
                for item in dd:
                    tag = " [FESTIVO]" if item.get("festivo") else ""
                    issues_str = "; ".join(item.get("issues", []))
                    lines.append(f"{item.get('date')} ({item.get('dow')}){tag}: {issues_str}")
        log_path.write_text("\n".join(lines), encoding="utf-8")
        return log_path
    except Exception:
        return None
def solve_across_months(
    cfg: dict,
    days: List[DayRow],
    unav_map: Dict[str, Dict[dt.date, Set[str]]],
    carryover_by_month: Optional[dict] = None,
    fixed_assignments: Optional[List[dict]] = None,
    availability_preferences: Optional[List[dict]] = None,
    v_double_overrides: Optional[List[str]] = None,
    j_blank_week_overrides: Optional[Dict[str, List[str]]] = None,
    disabled_columns: Optional[List[str]] = None,
    historical_stats: Optional[dict] = None,
    prior_usage: Optional[dict] = None,
) -> Tuple[List[Slot], Dict[str, Optional[str]], Dict]:
    """Solve schedules month-by-month and merge.

    `carryover_by_month` lets the caller inject cross-month constraints/history.
    `fixed_assignments`: list of {"doctor": str, "date": "YYYY-MM-DD", "column": str}
        Il medico DEVE comparire in quella colonna quel giorno (vincolo hard).
    `availability_preferences`: list of {"doctor": str, "date": "YYYY-MM-DD", "shift": str}
        Il software PROVA (soft) a far comparire il medico nella fascia indicata.
    """
    month_keys = sorted({(d.date.year, d.date.month) for d in days})
    slots_all: List[Slot] = []
    assignment_all: Dict[str, Optional[str]] = {}
    stats_all: Dict = {"status": "OK", "months": {}}

    gc = (cfg.get("global_constraints") or {})
    night_gap = int(gc.get("night_spacing_days_min", 5) or 5)
    night_off = (gc.get("night_off") or {}) if isinstance(gc.get("night_off"), dict) else {}
    night_off_next_day = bool(night_off.get("next_day", True))

    def _norm_key(yy: int, mm: int) -> str:
        return f"{yy}-{mm:02d}"

    def _parse_iso_date(s: str) -> Optional[dt.date]:
        try:
            return dt.date.fromisoformat(str(s).strip())
        except Exception:
            return None

    def _add_unav(local: Dict[str, Dict[dt.date, Set[str]]], doc: str, day: dt.date, shifts: Set[str]) -> None:
        docn = norm_name(doc)
        if not docn:
            return
        if docn not in local:
            local[docn] = {}
        if day not in local[docn]:
            local[docn][day] = set()
        local[docn][day].update(shifts)

    generated_carryover_by_month: Dict[str, dict] = {}
    running_prior_usage = deepcopy(prior_usage or {})
    running_prior_usage.setdefault("period_counts", {})

    for _month_idx, (yy, mm) in enumerate(month_keys):
        days_m = [d for d in days if (d.date.year, d.date.month) == (yy, mm)]
        mk = _norm_key(yy, mm)

        # Local copy of unavailability so we can inject carryover constraints without mutating input
        local_unav: Dict[str, Dict[dt.date, Set[str]]] = {k: {dk: set(sv) for dk, sv in v.items()} for k, v in (unav_map or {}).items()}

        carry_parts = []
        if generated_carryover_by_month.get(mk):
            carry_parts.append(generated_carryover_by_month.get(mk))
        if isinstance(carryover_by_month, dict) and carryover_by_month.get(mk):
            carry_parts.append(carryover_by_month.get(mk))
        carry = None
        prior_month_nights = ((prior_usage or {}).get("night_dates_by_doc") or {}).get(mk) or {}
        if prior_month_nights and days_m:
            prior_part = {"blocked_day1_doctors": [], "recent_nights_by_doc": {}}
            first_day_m = min(d.date for d in days_m)
            for dname, date_list in prior_month_nights.items():
                for ds in (date_list or []):
                    prev = _parse_iso_date(ds)
                    if not prev or prev >= first_day_m:
                        continue
                    delta = (first_day_m - prev).days
                    if 0 < delta < night_gap:
                        prior_part["recent_nights_by_doc"].setdefault(dname, []).append(prev.isoformat())
                    if delta == 1:
                        prior_part["blocked_day1_doctors"].append(dname)
            if prior_part["blocked_day1_doctors"] or prior_part["recent_nights_by_doc"]:
                carry_parts.append(prior_part)

        if carry_parts:
            carry = {"blocked_day1_doctors": [], "recent_nights_by_doc": {}}
            for part in carry_parts:
                for dname in (part.get("blocked_day1_doctors") or []):
                    if dname not in carry["blocked_day1_doctors"]:
                        carry["blocked_day1_doctors"].append(dname)
                recent_part = part.get("recent_nights_by_doc") or {}
                if isinstance(recent_part, dict):
                    for dname, date_list in recent_part.items():
                        carry["recent_nights_by_doc"].setdefault(dname, [])
                        for ds in (date_list or []):
                            if ds not in carry["recent_nights_by_doc"][dname]:
                                carry["recent_nights_by_doc"][dname].append(ds)
        if carry and days_m:
            # Block day 1 entirely for doctors coming from previous-month last night
            for dname in (carry.get("blocked_day1_doctors") or []):
                _add_unav(local_unav, dname, days_m[0].date, {"Any"})

            # Enforce night spacing at the beginning of the month
            recent = carry.get("recent_nights_by_doc") or {}
            if isinstance(recent, dict):
                for dname, date_list in recent.items():
                    for ds in (date_list or []):
                        prev = _parse_iso_date(ds)
                        if not prev:
                            continue
                        for drow in days_m:
                            delta = (drow.date - prev).days
                            if 0 <= delta < night_gap:
                                _add_unav(local_unav, dname, drow.date, {"Notte"})
                            if night_off_next_day and delta == 1:
                                _add_unav(local_unav, dname, drow.date, {"Mattina", "Pomeriggio"})

        # Filtra fixed_assignments e availability_preferences per questo mese
        fixed_m = []
        for f in (fixed_assignments or []):
            d = _parse_iso_date(str(f.get("date", "")))
            if d is not None and d.year == yy and d.month == mm:
                fixed_m.append(f)

        slots_m = slots_for_month(
            cfg,
            days_m,
            local_unav,
            fixed_assignments=fixed_m,
            v_double_overrides=v_double_overrides,
            j_blank_week_overrides=j_blank_week_overrides,
            disabled_columns=disabled_columns,
        )
        avail_m = []
        for a in (availability_preferences or []):
            d = _parse_iso_date(str(a.get("date", "")))
            if d is not None and d.year == yy and d.month == mm:
                avail_m.append(a)

        try:
            assignment_m, stats_m = solve_with_ortools(
                cfg, days_m, slots_m,
                fixed_assignments=fixed_m,
                availability_preferences=avail_m,
                unav_map=local_unav,
                historical_stats=historical_stats,
                prior_usage=running_prior_usage,
            )
        except Exception as e:
            # MODIFICA 2: niente Greedy. Se OR-Tools fallisce, propaga l'errore
            # e lascia le celle vuote (gestite dalla logica blank_penalty già esistente).
            stats_m = {
                "status": "INFEASIBLE",
                "solver_error": str(e),
                "note": "Greedy disabilitato: turni non coperti rimarranno vuoti (celle gialle).",
            }
            assignment_m = {}


        # Track where relief valves / blanks were used (useful for logs)
        try:
            stats_m = dict(stats_m or {})
            if j_blank_week_overrides:
                month_j_overrides: Dict[str, List[str]] = {}
                month_iso_weeks = {d.date.isocalendar()[:2] for d in days_m}
                for wk_key, vals in (j_blank_week_overrides or {}).items():
                    wk_tuple = None
                    try:
                        parts = str(wk_key).split("-W")
                        wk_tuple = (int(parts[0]), int(parts[1]))
                    except Exception:
                        pass
                    vals_in_month = []
                    for raw in vals or []:
                        dd = _parse_iso_date(str(raw))
                        if dd is not None and dd.year == yy and dd.month == mm:
                            vals_in_month.append(dd.isoformat())
                    if vals_in_month or wk_tuple in month_iso_weeks:
                        month_j_overrides[str(wk_key)] = sorted(vals_in_month)
                if month_j_overrides:
                    stats_m["j_blank_week_overrides"] = month_j_overrides
            stats_m["relief_used"] = build_relief_log(days_m, slots_m, assignment_m)
        except Exception:
            pass
        # Day-by-day diagnostic
        try:
            stats_m["daily_diagnostic"] = build_daily_diagnostic(days_m, slots_m, assignment_m, cfg)
        except Exception:
            pass

        # Merge
        slots_all.extend(slots_m)
        assignment_all.update(assignment_m)
        stats_all["months"][mk] = stats_m
        try:
            period_counts = running_prior_usage.setdefault("period_counts", {})
            for s in slots_m:
                doc = assignment_m.get(s.slot_id)
                if not doc:
                    continue
                doc_n = norm_name(doc)
                if s.rule_tag in ("Festivo_DE", "Festivo_HI"):
                    period_counts.setdefault(doc_n, {})
                    period_counts[doc_n]["Festivi"] = int(period_counts[doc_n].get("Festivi", 0) or 0) + 1
                elif s.columns == ["J"]:
                    period_counts.setdefault(doc_n, {})
                    period_counts[doc_n]["J"] = int(period_counts[doc_n].get("J", 0) or 0) + 1
                    if is_festivo(s.day, cfg) or s.day.date.weekday() == 5:
                        period_counts[doc_n]["J_FESTIVI"] = int(
                            period_counts[doc_n].get("J_FESTIVI", 0) or 0
                        ) + 1
        except Exception:
            pass
        st = str(stats_m.get("status", "")).upper()
        if "INFEAS" in st:
            stats_all["status"] = "INFEASIBLE"
        elif "PARTIAL" in st:
            stats_all["status"] = "PARTIAL"
        elif stats_all.get("status") == "OK" and "FEAS" in st:
            stats_all["status"] = "FEASIBLE"

        # If this generated range spans months, carry the solved end-of-month
        # nights into the next month before building its slots.
        if _month_idx + 1 < len(month_keys) and days_m:
            next_yy, next_mm = month_keys[_month_idx + 1]
            next_days = [d for d in days if (d.date.year, d.date.month) == (next_yy, next_mm)]
            if next_days:
                next_mk = _norm_key(next_yy, next_mm)
                next_first = min(d.date for d in next_days)
                last_date = max(d.date for d in days_m)
                recent_by_doc: Dict[str, List[str]] = {}
                last_night_doc = None
                for s in slots_m:
                    if s.columns != ["J"]:
                        continue
                    doc = assignment_m.get(s.slot_id)
                    if not doc:
                        continue
                    delta = (next_first - s.day.date).days
                    if 0 <= delta < night_gap:
                        recent_by_doc.setdefault(doc, []).append(s.day.date.isoformat())
                    if s.day.date == last_date:
                        last_night_doc = doc
                if recent_by_doc or last_night_doc:
                    generated_carryover_by_month[next_mk] = {
                        "blocked_day1_doctors": [last_night_doc] if last_night_doc else [],
                        "recent_nights_by_doc": recent_by_doc,
                    }

    try:
        repaired_assignment, swaps = _repair_heavy_festive_assignments(
            cfg,
            slots_all,
            assignment_all,
            fixed_assignments=fixed_assignments,
        )
        if swaps:
            assignment_all = repaired_assignment
            stats_all["heavy_festive_second_look"] = swaps
            slots_by_month: Dict[str, List[Slot]] = defaultdict(list)
            days_by_month: Dict[str, List[DayRow]] = defaultdict(list)
            for s in slots_all:
                slots_by_month[_norm_key(s.day.date.year, s.day.date.month)].append(s)
            for d in days:
                days_by_month[_norm_key(d.date.year, d.date.month)].append(d)
            for mk, sm in (stats_all.get("months") or {}).items():
                try:
                    sm["relief_used"] = build_relief_log(days_by_month.get(mk, []), slots_by_month.get(mk, []), assignment_all)
                    sm["daily_diagnostic"] = build_daily_diagnostic(days_by_month.get(mk, []), slots_by_month.get(mk, []), assignment_all, cfg)
                except Exception:
                    pass
    except Exception as e:
        stats_all.setdefault("warnings", []).append(f"second-look pesanti non applicato: {e}")

    return slots_all, assignment_all, stats_all


def extract_carryover_from_output_xlsx(
    output_xlsx: "Path | str",
    sheet_name: Optional[str] = None,
    night_col_letter: str = "J",
    min_gap: int = 5,
) -> dict:
    """Extract carryover information from a previously generated output Excel.

    Returns:
      {
        "source_last_date": "YYYY-MM-DD",
        "night_last_day_doctor": "Name" or None,
        "blocked_day1_doctors": ["Name"] (0/1 element),
        "recent_nights_by_doc": {"Name": ["YYYY-MM-DD", ...]}
      }
    """
    p = Path(output_xlsx)
    wb = openpyxl.load_workbook(p, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    # Assume Date in column A with header in row 1
    dates = []
    night_docs = []

    col_night = openpyxl.utils.column_index_from_string(night_col_letter)

    for r in range(2, ws.max_row + 1):
        dv = ws.cell(r, 1).value
        if dv is None:
            continue
        # Convert to date
        if isinstance(dv, dt.datetime):
            d = dv.date()
        elif isinstance(dv, dt.date):
            d = dv
        else:
            try:
                # strings like 2026-02-01
                d = dt.date.fromisoformat(str(dv)[:10])
            except Exception:
                continue
        nd = ws.cell(r, col_night).value
        nd = norm_name(nd) if nd else None
        dates.append(d)
        night_docs.append(nd)

    if not dates:
        return {
            "source_last_date": None,
            "night_last_day_doctor": None,
            "blocked_day1_doctors": [],
            "recent_nights_by_doc": {},
        }

    # sort by date just in case
    combined = sorted(zip(dates, night_docs), key=lambda x: x[0])
    dates, night_docs = zip(*combined)
    last_date = dates[-1]
    last_night_doc = night_docs[-1]

    # recent window: last (min_gap-1) days (inclusive of last day)
    recent_n = max(int(min_gap) - 1, 0)
    window = list(zip(dates[-recent_n:] if recent_n else [], night_docs[-recent_n:] if recent_n else []))

    recent_by_doc: Dict[str, List[str]] = {}
    for d, doc in window:
        if not doc:
            continue
        recent_by_doc.setdefault(doc, []).append(d.isoformat())

    return {
        "source_last_date": last_date.isoformat(),
        "night_last_day_doctor": last_night_doc,
        "blocked_day1_doctors": [last_night_doc] if last_night_doc else [],
        "recent_nights_by_doc": recent_by_doc,
    }


def generate_schedule(
    template_xlsx: "Path | str",
    rules_yml: "Path | str",
    out_xlsx: "Path | str",
    unavailability_path: "Path | str | None" = None,
    sheet_name: "str | None" = None,
    carryover_by_month: Optional[dict] = None,
    fixed_assignments: Optional[List[dict]] = None,
    availability_preferences: Optional[List[dict]] = None,
    v_double_overrides: Optional[List[str]] = None,
    j_blank_week_overrides: Optional[Dict[str, List[str]]] = None,
    disabled_columns: Optional[List[str]] = None,
    historical_stats: Optional[dict] = None,
    pool_config: Optional[dict] = None,
    prior_usage: Optional[dict] = None,
):
    """Generate schedules without Tkinter.

    This is the function used by Streamlit (and can be used programmatically).
    fixed_assignments: [{"doctor":str,"date":"YYYY-MM-DD","column":str}, ...]
    availability_preferences: [{"doctor":str,"date":"YYYY-MM-DD","shift":str}, ...]
    pool_config: dict caricato da pool_config_store (overlay JSON su YAML).
    """
    template = Path(template_xlsx)
    rules = Path(rules_yml)
    outp = Path(out_xlsx)
    unav = Path(unavailability_path) if unavailability_path else None

    cfg = load_rules(rules)
    if pool_config:
        cfg = apply_pool_config(cfg, pool_config)
    if disabled_columns:
        cfg["disabled_columns"] = sorted({
            str(c).strip().upper()
            for c in disabled_columns
            if str(c).strip()
        })

    # Merge pre-assigned holiday shifts from data/turni_festivi.yml
    tf = load_turni_festivi()
    if tf["festivi_extra"]:
        existing_fe = list(cfg.get("festivi_extra") or [])
        existing_fe_set = set(str(x).strip() for x in existing_fe)
        cfg["festivi_extra"] = existing_fe + [x for x in tf["festivi_extra"] if x not in existing_fe_set]
    if tf["fixed_assignments"]:
        fixed_assignments = list(fixed_assignments or []) + tf["fixed_assignments"]

    wb, ws, days = load_template_days(template, sheet_name=sheet_name)
    unav_map = load_unavailability(unav)
    # Strip unavailability entries that conflict with pre-assigned festivo sorteggio shifts
    if tf["fixed_assignments"]:
        _strip_festivi_unavailability(unav_map, tf["fixed_assignments"])
    slots, assignment, stats = solve_across_months(
        cfg, days, unav_map,
        carryover_by_month=carryover_by_month,
        v_double_overrides=v_double_overrides,
        j_blank_week_overrides=j_blank_week_overrides,
        disabled_columns=disabled_columns,
        fixed_assignments=fixed_assignments,
        availability_preferences=availability_preferences,
        historical_stats=historical_stats,
        prior_usage=prior_usage,
    )
    # REMOVED: terza chiamata ridondante a assign_reperibilita_C (sovrascriveva C già ottimizzata)
    write_output(wb, ws, days, slots, assignment, cfg=cfg, out_path=outp, unav_map=unav_map)
    logp = write_solver_log(outp, stats)
    return stats, str(logp) if logp else None

# GUI (tkinter)
# -------------------------
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.title("Turni Autogenerator (prototype)")
    template_var = tk.StringVar()
    rules_var = tk.StringVar()
    unav_var = tk.StringVar()
    out_var = tk.StringVar()
    def pick_file(var: tk.StringVar, types):
        p = filedialog.askopenfilename(filetypes=types)
        if p:
            var.set(p)
    def pick_save(var: tk.StringVar):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            var.set(p)
    def go():
        try:
            template = Path(template_var.get())
            rules = Path(rules_var.get())
            unav = Path(unav_var.get()) if unav_var.get().strip() else None
            outp = Path(out_var.get())
            if not template.exists() or not rules.exists() or not outp:
                raise ValueError("Seleziona template, regole e output.")
            cfg = load_rules(rules)
            wb, ws, days = load_template_days(template)
            unav_map = load_unavailability(unav)
            slots, assignment, stats = solve_across_months(cfg, days, unav_map)
            # Avviso se qualche mese è andato in fallback greedy
            greedy_months = [k for k,v in (stats.get('months') or {}).items() if isinstance(v, dict) and v.get('solver_error')]
            if greedy_months:
                messagebox.showwarning(
                    "Solver",
                    "OR-Tools non disponibile o schedule infeasible per: " + ", ".join(greedy_months) + "\nUso greedy per quei mesi."
                )
            write_output(wb, ws, days, slots, assignment, cfg=cfg, out_path=outp, unav_map=unav_map)
            logp = write_solver_log(outp, stats)
            msg = f"Creato: {outp}\nSolver: {stats.get('status')}"
            if logp:
                msg += f"\nLog: {logp}"
            messagebox.showinfo("OK", msg)
        except Exception as e:
            messagebox.showerror("Errore", str(e))
    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack(fill="both", expand=True)
    def row(label, var, btn_text, cmd, r):
        tk.Label(frm, text=label, anchor="w").grid(row=r, column=0, sticky="w")
        tk.Entry(frm, textvariable=var, width=60).grid(row=r, column=1, padx=5)
        tk.Button(frm, text=btn_text, command=cmd).grid(row=r, column=2)
    row("Template turni (.xlsx)", template_var, "Scegli...", lambda: pick_file(template_var,[("Excel","*.xlsx")]), 0)
    row("Regole (.yml)", rules_var, "Scegli...", lambda: pick_file(rules_var,[("YAML","*.yml;*.yaml")]), 1)
    row("Indisponibilità (opz.)", unav_var, "Scegli...", lambda: pick_file(unav_var,[("Excel/CSV","*.xlsx;*.xls;*.csv;*.tsv")]), 2)
    row("Output (.xlsx)", out_var, "Salva come...", lambda: pick_save(out_var), 3)
    tk.Button(frm, text="Genera turni", command=go, height=2).grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")
    root.mainloop()
# -------------------------
# Main
# -------------------------
def main():
    ap = argparse.ArgumentParser(description="Turni Autogenerator (UTIC/Cardiologia)")
    ap.add_argument("--template", type=str, help="Template Excel .xlsx")
    ap.add_argument("--rules", type=str, help="Regole YAML .yml/.yaml")
    ap.add_argument("--unavailability", type=str, default="", help="Indisponibilità mensili .xlsx/.csv (opzionale)")
    ap.add_argument("--out", type=str, help="Output Excel .xlsx")
    ap.add_argument("--sheet", type=str, default="", help="Nome foglio (opzionale)")
    ap.add_argument("--gui", action="store_true", help="Avvia interfaccia grafica (Tkinter)")
    args = ap.parse_args()
    if args.gui or (not args.template and not args.rules):
        run_gui()
        return
    if not args.template or not args.rules or not args.out:
        ap.error("In modalità CLI devi specificare: --template, --rules, --out")
    template = Path(args.template)
    rules = Path(args.rules)
    outp = Path(args.out)
    unav = Path(args.unavailability) if args.unavailability.strip() else None
    cfg = load_rules(rules)
    wb, ws, days = load_template_days(template, sheet_name=args.sheet if args.sheet else None)
    unav_map = load_unavailability(unav)
    slots, assignment, stats = solve_across_months(cfg, days, unav_map)
    # If any month fell back to greedy, print it
    greedy_months = [k for k,v in (stats.get('months') or {}).items() if isinstance(v, dict) and v.get('solver_error')]
    if greedy_months:
        print("[WARN] OR-Tools non disponibile o infeasible per:", ", ".join(greedy_months), file=sys.stderr)
    write_output(wb, ws, days, slots, assignment, cfg=cfg, out_path=outp, unav_map=unav_map)
    logp = write_solver_log(outp, stats)
    if logp:
        print(f"OK: creato {outp} | solver={stats.get('status')} | log={logp}")
    else:
        print(f"OK: creato {outp} | solver={stats.get('status')}")
if __name__ == "__main__":
    main()
