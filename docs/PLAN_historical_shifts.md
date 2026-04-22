# Memoria Storica Turni — Piano di Implementazione

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Permettere all'admin di caricare i file Excel definitivi dei mesi precedenti, costruire una memoria storica dei turni per medico, e usarla nel solver per bilanciare quote cross-mese (notti, reperibilità, festivi, domeniche). Dashboard admin con grafici e tabelle.

**Architecture:** Nuovo modulo `shift_history.py` che: (1) parsa i file Excel definitivi estraendo le assegnazioni per medico/giorno/colonna, (2) salva lo storico come JSON su GitHub (stesso meccanismo di `github_utils.py`), (3) calcola aggregazioni cross-mese. Il solver riceve un dict `historical_stats` con i conteggi cumulativi e li usa come soft-constraints per bilanciare. La UI admin aggiunge una sezione "Memoria Storica" con upload, tabelle e grafici Plotly.

**Tech Stack:** Python 3.11, openpyxl (parsing Excel), Streamlit (UI), Plotly (grafici), OR-Tools CP-SAT (soft constraints), GitHub API (storage JSON)

---

## File Structure

| File | Responsabilità | Azione |
|------|---------------|--------|
| `shift_history.py` | Parser Excel definitivo → estrazione assegnazioni; aggregazione stats; serializzazione JSON | **NUOVO** |
| `turni_generator.py` | Riceve `historical_stats` nel solver, aggiunge soft-constraints per bilanciamento cross-mese | **MODIFICA** (linee ~2260-2400 quota section + ~3751 solve_across_months + ~3963 generate_schedule) |
| `streamlit_app.py` | Sezione admin "Memoria Storica": upload file, tabella riepilogo, grafici, integrazione nel flusso di generazione | **MODIFICA** (dopo sezione carryover ~2970) |
| `github_utils.py` | Nessuna modifica — riusato as-is per read/write JSON su GitHub | Invariato |
| `Regole_Turni.yml` | Nessuna modifica strutturale — le quote fisse restano; il solver le rispetta e bilancia solo le quote "libere" | Invariato |

---

## Task 1: Parser Excel definitivo — `shift_history.py`

**Files:**
- Create: `shift_history.py`

**Contesto:** Il file Excel definitivo (es. `Turni Aprile 2026.xlsx`) ha:
- Colonna A = date, Colonna B = giorno settimana
- Colonne C-AC = assegnazioni medici (un nome per cella, o due separati da `\n` in V/Y)
- Colonne M-N-O-P = emodinamica (possono contenere nomi medici aggiunti a mano dal primario)
- Foglio "Riepilogo" = riepilogo automatico (lo ignoriamo, parsiamo il foglio principale)

- [ ] **Step 1: Creare `shift_history.py` con la funzione `parse_finalized_xlsx()`**

```python
# shift_history.py
"""Memoria storica turni: parsing file definitivi e aggregazione cross-mese."""
from __future__ import annotations

import datetime as dt
import json
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

# Colonne da tracciare per lo storico.
# Chiave = lettera colonna Excel, valore = etichetta umana.
TRACKED_COLUMNS = {
    "C": "Reperibilità",
    "D": "UTIC mattina",
    "E": "Cardiologia mattina",
    "F": "Supporto 118",
    "G": "Riabilitazione",
    "H": "UTIC pomeriggio",
    "I": "Cardiologia pomeriggio",
    "J": "Notte",
    "K": "Letto",
    "L": "Padiglioni",
    "M": "Emodinamica 1",
    "N": "Emodinamica 2",
    "O": "Emodinamica 3",
    "P": "Emodinamica 4",
    "Q": "ECO base",
    "R": "ECOSTRESS/ETE",
    "S": "Ecosala",
    "T": "Interni",
    "U": "Contr.PM",
    "V": "Sala PM",
    "W": "Ergometria/CPET",
    "Y": "Ambulatori",
    "Z": "Vascolare",
    "AA": "SPOC",
    "AB": "Holter/Brugada/FA",
    "AC": "Scintigrafia",
}

# Giorni festivi italiani (mese, giorno)
_FESTIVI_FISSI = {
    (1, 1), (1, 6), (4, 25), (5, 1), (6, 2),
    (8, 15), (11, 1), (12, 8), (12, 25), (12, 26),
}


def _is_holiday(d: dt.date) -> bool:
    """True se domenica o festivo nazionale."""
    if d.weekday() == 6:  # domenica
        return True
    return (d.month, d.day) in _FESTIVI_FISSI


def _norm(name: Optional[str]) -> Optional[str]:
    """Normalizza un nome medico (strip + title case)."""
    if not name:
        return None
    s = str(name).strip()
    if not s or s.lower() in ("", "none", "nan"):
        return None
    return s


def _extract_doctors_from_cell(value) -> List[str]:
    """Estrae uno o più nomi medico da una cella (gestisce newline per V/Y)."""
    if value is None:
        return []
    s = str(value).strip()
    if not s or s.lower() in ("none", "nan"):
        return []
    # Celle con più medici separati da newline (colonne V, Y)
    parts = s.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    result = []
    for p in parts:
        name = _norm(p)
        if name and name != "Recupero":
            result.append(name)
    return result


def parse_finalized_xlsx(
    xlsx_path: "Path | str",
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Parsa un file Excel definitivo e restituisce le assegnazioni.

    Returns:
        {
            "year": 2026,
            "month": 4,
            "month_label": "2026-04",
            "days": [
                {
                    "date": "2026-04-01",
                    "dow": "Wed",
                    "is_holiday": false,
                    "assignments": {
                        "C": ["Licordari"],
                        "D": ["Grimaldi"],
                        "J": ["Crea"],
                        "V": ["Crea", "Dattilo"],  # multi-doctor
                        ...
                    }
                },
                ...
            ]
        }
    """
    wb = openpyxl.load_workbook(Path(xlsx_path), data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    col_indices = {}
    for col_letter in TRACKED_COLUMNS:
        col_indices[col_letter] = openpyxl.utils.column_index_from_string(col_letter)

    days_data = []
    year = None
    month = None

    for r in range(2, ws.max_row + 1):
        dv = ws.cell(r, 1).value
        if dv is None:
            continue
        if isinstance(dv, dt.datetime):
            d = dv.date()
        elif isinstance(dv, dt.date):
            d = dv
        else:
            try:
                d = dt.date.fromisoformat(str(dv)[:10])
            except Exception:
                continue

        if year is None:
            year = d.year
            month = d.month

        assignments = {}
        for col_letter, col_idx in col_indices.items():
            cell_val = ws.cell(r, col_idx).value
            doctors = _extract_doctors_from_cell(cell_val)
            if doctors:
                assignments[col_letter] = doctors

        dow_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        days_data.append({
            "date": d.isoformat(),
            "dow": dow_names[d.weekday()],
            "is_holiday": _is_holiday(d),
            "assignments": assignments,
        })

    wb.close()

    return {
        "year": year or 0,
        "month": month or 0,
        "month_label": f"{year or 0}-{(month or 0):02d}",
        "days": days_data,
    }


def compute_doctor_stats(parsed: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Calcola le statistiche per medico da un mese parsato.

    Returns:
        {
            "Licordari": {
                "C": {"total": 3, "festivi": 0, "feriali": 3},
                "J": {"total": 3, "festivi": 1, "feriali": 2, "sabati": 0, "domeniche": 1},
                "H": {"total": 4, "festivi": 1, "feriali": 3},
                ...
                "_festivi_DE_HI": 2,  # turni festivi mattina+pomeriggio (D/E/H/I in festivi)
                "_domeniche": 1,      # domeniche lavorate (qualsiasi colonna attiva)
                "_sabati": 0,
            },
            ...
        }
    """
    stats: Dict[str, Dict[str, Any]] = {}

    for day_data in parsed.get("days", []):
        d = dt.date.fromisoformat(day_data["date"])
        is_hol = day_data.get("is_holiday", False)
        is_sun = d.weekday() == 6
        is_sat = d.weekday() == 5

        for col, doctors in day_data.get("assignments", {}).items():
            for doc in doctors:
                if doc not in stats:
                    stats[doc] = {}

                if col not in stats[doc]:
                    stats[doc][col] = {"total": 0, "festivi": 0, "feriali": 0}

                stats[doc][col]["total"] += 1
                if is_hol:
                    stats[doc][col]["festivi"] += 1
                else:
                    stats[doc][col]["feriali"] += 1

                # Per colonna J: traccia sabati e domeniche separatamente
                if col == "J":
                    stats[doc][col].setdefault("sabati", 0)
                    stats[doc][col].setdefault("domeniche", 0)
                    if is_sat:
                        stats[doc][col]["sabati"] += 1
                    if is_sun:
                        stats[doc][col]["domeniche"] += 1

        # Conteggi aggregati festivi/weekend
        if is_hol:
            fest_cols = {"D", "E", "H", "I"}
            for col in fest_cols:
                for doc in day_data.get("assignments", {}).get(col, []):
                    stats.setdefault(doc, {})
                    stats[doc].setdefault("_festivi_DE_HI", 0)
                    stats[doc]["_festivi_DE_HI"] += 1

        if is_sun:
            # Colonne "attive" (esclusa C che è passiva)
            active_cols = set(day_data.get("assignments", {}).keys()) - {"C"}
            worked_docs = set()
            for col in active_cols:
                for doc in day_data.get("assignments", {}).get(col, []):
                    worked_docs.add(doc)
            for doc in worked_docs:
                stats.setdefault(doc, {})
                stats[doc].setdefault("_domeniche", 0)
                stats[doc]["_domeniche"] += 1

        if is_sat:
            active_cols = set(day_data.get("assignments", {}).keys()) - {"C"}
            worked_docs = set()
            for col in active_cols:
                for doc in day_data.get("assignments", {}).get(col, []):
                    worked_docs.add(doc)
            for doc in worked_docs:
                stats.setdefault(doc, {})
                stats[doc].setdefault("_sabati", 0)
                stats[doc]["_sabati"] += 1

    return stats


def aggregate_multi_month(
    all_months: Dict[str, Dict[str, Dict[str, Any]]],
) -> Dict[str, Dict[str, Any]]:
    """Aggrega le stats di più mesi in un unico riepilogo cumulativo.

    Args:
        all_months: {"2026-03": {doctor_stats}, "2026-04": {doctor_stats}, ...}

    Returns:
        {
            "Licordari": {
                "C": {"total": 7, "festivi": 1, "feriali": 6},
                "J": {"total": 6, "festivi": 2, "feriali": 4, "sabati": 1, "domeniche": 1},
                "_festivi_DE_HI": 4,
                "_domeniche": 3,
                "_sabati": 1,
                "_months_counted": 2,
            },
            ...
        }
    """
    agg: Dict[str, Dict[str, Any]] = {}

    for month_label, month_stats in sorted(all_months.items()):
        for doc, doc_data in month_stats.items():
            if doc not in agg:
                agg[doc] = {"_months_counted": 0}

            for key, val in doc_data.items():
                if key.startswith("_"):
                    # Scalar counters (_festivi_DE_HI, _domeniche, _sabati)
                    agg[doc][key] = agg[doc].get(key, 0) + val
                else:
                    # Column stats dict
                    if key not in agg[doc]:
                        agg[doc][key] = {"total": 0, "festivi": 0, "feriali": 0}
                    for subkey, subval in val.items():
                        agg[doc][key][subkey] = agg[doc][key].get(subkey, 0) + subval

            agg[doc]["_months_counted"] = agg[doc].get("_months_counted", 0) + 1

    return agg
```

- [ ] **Step 2: Aggiungere funzioni di serializzazione JSON per GitHub**

Aggiungere in fondo a `shift_history.py`:

```python
def history_to_json(
    history: Dict[str, Dict[str, Dict[str, Any]]],
) -> str:
    """Serializza lo storico completo (tutti i mesi) in JSON.

    Args:
        history: {"2026-03": {doctor_stats}, "2026-04": {doctor_stats}, ...}
    """
    return json.dumps(history, ensure_ascii=False, indent=2, sort_keys=True)


def history_from_json(text: str) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """Deserializza lo storico da JSON."""
    if not text or not text.strip():
        return {}
    return json.loads(text)
```

- [ ] **Step 3: Testare il parser con il file Aprile 2026**

```bash
cd "C:\Users\Roberto Licordari\Desktop\turni_autogen"
python -c "
import shift_history as sh
parsed = sh.parse_finalized_xlsx('data/Turni Aprile 2026.xlsx')
print(f'Mese: {parsed[\"month_label\"]}, Giorni: {len(parsed[\"days\"])}')
stats = sh.compute_doctor_stats(parsed)
for doc in sorted(stats):
    j = stats[doc].get('J', {})
    c = stats[doc].get('C', {})
    dom = stats[doc].get('_domeniche', 0)
    print(f'  {doc}: J={j.get(\"total\",0)} (fest={j.get(\"festivi\",0)}, sab={j.get(\"sabati\",0)}, dom={j.get(\"domeniche\",0)}), C={c.get(\"total\",0)}, domeniche={dom}')
"
```

Verificare che i conteggi corrispondano al Riepilogo del file Excel.

- [ ] **Step 4: Commit**

```bash
git add shift_history.py
git commit -m "feat: nuovo modulo shift_history — parser Excel definitivo e aggregazione stats"
```

---

## Task 2: Storage su GitHub — lettura/scrittura storico

**Files:**
- Modify: `shift_history.py`

**Contesto:** Lo storico viene salvato come `data/shift_history.json` sulla stessa repo GitHub usata per le indisponibilità. Struttura:
```json
{
  "2026-03": { "Licordari": { "C": {"total": 2, ...}, "J": {...}, ... }, ... },
  "2026-04": { ... }
}
```

- [ ] **Step 1: Aggiungere funzioni load/save che usano `github_utils`**

Aggiungere in fondo a `shift_history.py`:

```python
from github_utils import get_file, put_file, GithubFile

HISTORY_PATH = "data/shift_history.json"


def load_history_from_github(
    owner: str, repo: str, token: str, branch: str = "main",
) -> Tuple[Dict[str, Dict[str, Dict[str, Any]]], Optional[str]]:
    """Carica lo storico da GitHub. Ritorna (history_dict, sha)."""
    gf = get_file(owner, repo, HISTORY_PATH, token, branch)
    if gf is None:
        return {}, None
    return history_from_json(gf.text), gf.sha


def save_history_to_github(
    history: Dict[str, Dict[str, Dict[str, Any]]],
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    sha: Optional[str] = None,
) -> dict:
    """Salva lo storico su GitHub."""
    text = history_to_json(history)
    return put_file(
        owner, repo, HISTORY_PATH, token,
        message=f"Aggiornamento memoria storica turni",
        text=text, branch=branch, sha=sha,
    )


def upload_month_to_history(
    xlsx_path: "Path | str",
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    sheet_name: Optional[str] = None,
) -> Tuple[str, Dict[str, Dict[str, Any]]]:
    """Pipeline completa: parsa xlsx → aggiorna storico su GitHub.

    Returns:
        (month_label, doctor_stats_for_that_month)
    """
    parsed = parse_finalized_xlsx(xlsx_path, sheet_name=sheet_name)
    month_label = parsed["month_label"]
    month_stats = compute_doctor_stats(parsed)

    # Load existing history
    history, sha = load_history_from_github(owner, repo, token, branch)

    # Upsert this month
    history[month_label] = month_stats

    # Save back
    save_history_to_github(history, owner, repo, token, branch, sha)

    return month_label, month_stats
```

- [ ] **Step 2: Commit**

```bash
git add shift_history.py
git commit -m "feat: shift_history load/save su GitHub via Contents API"
```

---

## Task 3: Integrazione nel solver — soft constraints con storico

**Files:**
- Modify: `turni_generator.py:3963-4009` (funzione `generate_schedule`)
- Modify: `turni_generator.py:3751-3760` (firma `solve_across_months`)
- Modify: `turni_generator.py:2257-2400` (sezione quota constraints in `solve_with_ortools`)

**Contesto:** Il solver deve ricevere un dict `historical_stats` (output di `aggregate_multi_month`) e usarlo come soft-constraints. Le quote fisse (Licordari 3 notti, De Gregorio 1 festivo, ecc.) restano HARD. Il bilanciamento storico agisce solo sulle quote "libere" — cioè i medici senza quota fissa.

Strategia di bilanciamento:
- **Notti (J) — chi fa la terza notte:** Tra i medici senza quota fissa (min=2, max=3), il solver preferisce assegnare 3 notti a chi storicamente ne ha fatte meno. Penalità proporzionale al conteggio storico.
- **Domeniche (Festivi DE/HI):** Tra i medici senza quota fissa per festivi, il solver preferisce assegnare a chi ha meno domeniche cumulate.
- **Reperibilità (C):** Il target è 2-3 per mese; il solver preferisce dare 3 a chi ne ha fatte meno.
- **H/I pomeriggio:** Già bilanciati internamente al mese; lo storico aggiunge un tie-breaker.

- [ ] **Step 1: Aggiungere parametro `historical_stats` a `generate_schedule` e `solve_across_months`**

In `turni_generator.py`, modificare la firma di `generate_schedule` (linea ~3963):

```python
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
    j_blank_week_overrides: Optional[Dict[str, Optional[str]]] = None,
    historical_stats: Optional[dict] = None,  # NUOVO
):
```

E passarlo a `solve_across_months`:

```python
    slots, assignment, stats = solve_across_months(
        cfg, days, unav_map,
        carryover_by_month=carryover_by_month,
        v_double_overrides=v_double_overrides,
        j_blank_week_overrides=j_blank_week_overrides,
        fixed_assignments=fixed_assignments,
        availability_preferences=availability_preferences,
        historical_stats=historical_stats,  # NUOVO
    )
```

Modificare la firma di `solve_across_months` (linea ~3751):

```python
def solve_across_months(
    cfg: dict,
    days: List[DayRow],
    unav_map: Dict[str, Dict[dt.date, Set[str]]],
    carryover_by_month: Optional[dict] = None,
    fixed_assignments: Optional[List[dict]] = None,
    availability_preferences: Optional[List[dict]] = None,
    v_double_overrides: Optional[List[str]] = None,
    j_blank_week_overrides: Optional[Dict[str, Optional[str]]] = None,
    historical_stats: Optional[dict] = None,  # NUOVO
) -> Tuple[List[Slot], Dict[str, Optional[str]], Dict]:
```

E passarlo a `solve_with_ortools` (linea ~3841):

```python
        try:
            assignment_m, stats_m = solve_with_ortools(
                cfg, days_m, slots_m,
                fixed_assignments=fixed_m,
                availability_preferences=avail_m,
                unav_map=local_unav,
                historical_stats=historical_stats,  # NUOVO
            )
```

- [ ] **Step 2: Aggiungere parametro `historical_stats` a `solve_with_ortools` e implementare soft constraints**

Modificare firma di `solve_with_ortools` per accettare `historical_stats: Optional[dict] = None`.

Dopo la sezione "Night distribution" (dopo linea ~2386), aggiungere il blocco di bilanciamento storico:

```python
    # ── Historical balance: soft constraints from previous months ──────
    hist = historical_stats or {}
    if hist:
        HIST_NIGHT_PENALTY = 150   # per notte storica in più rispetto alla media
        HIST_FEST_PENALTY = 200    # per festivo storico in più
        HIST_C_PENALTY = 100       # per reperibilità storica in più

        # --- J (Notti): chi tra i "free" ha fatto meno notti storiche → più probabilità di fare 3 ---
        if "rules" in cfg and "J" in cfg["rules"]:
            rJ = cfg["rules"]["J"]
            mq_fixed = {norm_name(k) for k in (rJ.get("monthly_quotas") or {}).keys()}
            night_pool_free = [d for d in sorted(night_pool) if d not in mq_fixed and d in doctors] if 'night_pool' in dir() else []

            if night_pool_free:
                hist_j_totals = {}
                for doc in night_pool_free:
                    j_data = hist.get(doc, {}).get("J", {})
                    hist_j_totals[doc] = j_data.get("total", 0) if isinstance(j_data, dict) else 0

                if any(v > 0 for v in hist_j_totals.values()):
                    for doc in night_pool_free:
                        vars_ = [night_var_by_day_doc.get((d.date, doc)) for d in days
                                 if night_var_by_day_doc.get((d.date, doc)) is not None]
                        if vars_:
                            cnt = model.NewIntVar(0, total_nights, f"hist_j_{hash(doc)%10**6}")
                            model.Add(cnt == sum(vars_))
                            # Penalità: storico * conteggio attuale (chi ha più storico paga di più per ogni notte)
                            extra_obj.append(HIST_NIGHT_PENALTY * hist_j_totals[doc] * cnt)

        # --- Festivi: chi ha meno domeniche storiche → preferito per festivi ---
        if "rules" in cfg and "Festivi" in cfg["rules"]:
            rFest = cfg["rules"]["Festivi"]
            fest_fixed = {norm_name(k) for k in (rFest.get("quotas") or {}).keys()}
            fest_pool_all = [norm_name(d) for d in (rFest.get("pool") or [])
                            if norm_name(d) in doctors and norm_name(d) != "Recupero"]
            fest_pool_free = [d for d in fest_pool_all if d not in fest_fixed]
            festivo_slots_hist = [s for s in slots if s.rule_tag in ("Festivo_DE", "Festivo_HI")]

            if fest_pool_free and festivo_slots_hist:
                for doc in fest_pool_free:
                    hist_fest = hist.get(doc, {}).get("_festivi_DE_HI", 0)
                    if isinstance(hist_fest, dict):
                        hist_fest = 0
                    vars_d = [x[(s.slot_id, doc)] for s in festivo_slots_hist
                              if (s.slot_id, doc) in x]
                    if vars_d and hist_fest > 0:
                        cnt = model.NewIntVar(0, len(festivo_slots_hist),
                                              f"hist_fest_{hash(doc)%10**6}")
                        model.Add(cnt == sum(vars_d))
                        extra_obj.append(HIST_FEST_PENALTY * hist_fest * cnt)

        # --- Domeniche J (notti): chi ha meno notti di domenica → preferito ---
        if "rules" in cfg and "J" in cfg["rules"]:
            for doc in [d for d in sorted(night_pool or set()) if d in doctors]:
                hist_dom_j = 0
                j_data = hist.get(doc, {}).get("J", {})
                if isinstance(j_data, dict):
                    hist_dom_j = j_data.get("domeniche", 0)
                if hist_dom_j > 0:
                    sun_vars = [night_var_by_day_doc.get((d.date, doc))
                                for d in days if d.dow == "Sun"]
                    sun_vars = [v for v in sun_vars if v is not None]
                    if sun_vars:
                        sun_cnt = model.NewIntVar(0, len(sun_vars), f"hist_sunj_{hash(doc)%10**6}")
                        model.Add(sun_cnt == sum(sun_vars))
                        extra_obj.append(HIST_FEST_PENALTY * hist_dom_j * sun_cnt)
```

**Nota:** Le penalità sono proporzionali allo storico: un medico con 6 notti cumulate nei mesi precedenti "costa" di più al solver rispetto a uno con 3 notti, quindi il solver preferirà il secondo. Questo non viola le quote fisse (che sono HARD) ma influenza la distribuzione delle quote variabili.

- [ ] **Step 3: Testare che il solver compili senza errori con `historical_stats={}`**

```bash
python -c "
import turni_generator as tg
stats, log = tg.generate_schedule(
    template_xlsx='data/Turni Aprile 2026.xlsx',
    rules_yml='Regole_Turni.yml',
    out_xlsx='test_output.xlsx',
    historical_stats={},
)
print('Status:', stats.get('status'))
import os; os.remove('test_output.xlsx') if os.path.exists('test_output.xlsx') else None
"
```

- [ ] **Step 4: Commit**

```bash
git add turni_generator.py
git commit -m "feat: solver accetta historical_stats per bilanciamento cross-mese"
```

---

## Task 4: Sezione admin "Memoria Storica" — Upload e Visualizzazione

**Files:**
- Modify: `streamlit_app.py` (dopo sezione carryover, prima di "Assegnazioni fisse")

**Contesto:** La sezione admin deve permettere:
1. Upload di un file Excel definitivo per un mese passato
2. Visualizzazione tabella riepilogo per medico (colonne chiave)
3. Grafici: notti per medico nel tempo, domeniche cumulate, reperibilità
4. Eliminazione di un mese dallo storico
5. Lo storico viene automaticamente passato a `generate_schedule` quando si generano i turni

- [ ] **Step 1: Aggiungere import di `shift_history` in cima a `streamlit_app.py`**

Aggiungere tra gli import esistenti:

```python
import shift_history as sh
```

- [ ] **Step 2: Aggiungere funzione helper per caricare/salvare lo storico via secrets**

Aggiungere come funzione helper (vicino alle altre funzioni helper dell'admin):

```python
def _load_shift_history() -> Tuple[dict, Optional[str]]:
    """Carica lo storico turni da GitHub."""
    try:
        sec = st.secrets["github_unavailability"]
        return sh.load_history_from_github(
            sec["owner"], sec["repo"], sec["token"], sec.get("branch", "main"),
        )
    except Exception:
        return {}, None


def _save_shift_history(history: dict, sha: Optional[str] = None) -> bool:
    """Salva lo storico turni su GitHub."""
    try:
        sec = st.secrets["github_unavailability"]
        sh.save_history_to_github(
            history, sec["owner"], sec["repo"], sec["token"],
            sec.get("branch", "main"), sha,
        )
        return True
    except Exception as e:
        st.error(f"Errore salvataggio storico: {e}")
        return False
```

- [ ] **Step 3: Inserire la sezione "Memoria Storica" nella UI admin**

Dopo la sezione carryover (~linea 2968) e prima di "Assegnazioni fisse" (~linea 2971), aggiungere:

```python
    st.divider()

    # ── Step 3b: Memoria Storica ────────────────────────────────────────────
    st.markdown("### 📊 Memoria Storica Turni")
    st.info(
        "Carica i file Excel **definitivi** dei mesi precedenti per costruire una memoria storica. "
        "Il solver userà questi dati per bilanciare le quote tra i mesi.",
        icon="🧠",
    )

    history_data, history_sha = _load_shift_history()

    # Upload new month
    with st.expander("📤 Carica mese definitivo", expanded=False):
        hist_upload = st.file_uploader(
            "File Excel turni definitivo",
            type=["xlsx"],
            key="hist_upload",
            help="Il file Excel finale (dopo le modifiche del primario) di un mese passato.",
        )
        if hist_upload is not None:
            if st.button("📥 Importa nel storico", key="btn_import_hist"):
                tmp_hist = Path(tempfile.gettempdir()) / f"hist_{int(time.time())}.xlsx"
                tmp_hist.write_bytes(hist_upload.getvalue())
                try:
                    parsed = sh.parse_finalized_xlsx(tmp_hist)
                    month_label = parsed["month_label"]
                    month_stats = sh.compute_doctor_stats(parsed)
                    history_data[month_label] = month_stats
                    if _save_shift_history(history_data, history_sha):
                        st.success(f"✅ Mese **{month_label}** importato ({len(parsed['days'])} giorni)")
                        st.rerun()
                except Exception as e:
                    st.error(f"Errore parsing: {e}")

    # Display loaded months
    if history_data:
        sorted_months = sorted(history_data.keys())
        st.caption(f"Mesi in memoria: {', '.join(sorted_months)}")

        # Aggregate stats
        agg = sh.aggregate_multi_month(history_data)

        # Table: key columns per doctor
        with st.expander("📋 Tabella riepilogativa", expanded=True):
            import pandas as pd
            rows = []
            for doc in sorted(agg.keys()):
                ds = agg[doc]
                rows.append({
                    "Medico": doc,
                    "Mesi": ds.get("_months_counted", 0),
                    "Notti (J)": ds.get("J", {}).get("total", 0) if isinstance(ds.get("J"), dict) else 0,
                    "Notti Sab": ds.get("J", {}).get("sabati", 0) if isinstance(ds.get("J"), dict) else 0,
                    "Notti Dom": ds.get("J", {}).get("domeniche", 0) if isinstance(ds.get("J"), dict) else 0,
                    "Reperibilità (C)": ds.get("C", {}).get("total", 0) if isinstance(ds.get("C"), dict) else 0,
                    "Festivi (D/E/H/I)": ds.get("_festivi_DE_HI", 0),
                    "Domeniche": ds.get("_domeniche", 0),
                    "Sabati": ds.get("_sabati", 0),
                    "H pom.": ds.get("H", {}).get("total", 0) if isinstance(ds.get("H"), dict) else 0,
                    "I pom.": ds.get("I", {}).get("total", 0) if isinstance(ds.get("I"), dict) else 0,
                })
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

        # Charts
        with st.expander("📈 Grafici", expanded=False):
            import plotly.express as px

            # Bar chart: notti per medico
            if rows:
                fig_j = px.bar(
                    df, x="Medico", y="Notti (J)",
                    title="Notti totali per medico (cumulativo)",
                    color="Notti (J)", color_continuous_scale="Reds",
                )
                fig_j.update_layout(xaxis_tickangle=-45, height=400)
                st.plotly_chart(fig_j, use_container_width=True)

                # Stacked bar: notti weekend vs feriali
                df_j_detail = pd.DataFrame([{
                    "Medico": r["Medico"],
                    "Feriali": r["Notti (J)"] - r["Notti Sab"] - r["Notti Dom"],
                    "Sabato": r["Notti Sab"],
                    "Domenica": r["Notti Dom"],
                } for r in rows])
                fig_j2 = px.bar(
                    df_j_detail, x="Medico", y=["Feriali", "Sabato", "Domenica"],
                    title="Notti: distribuzione feriali/sabato/domenica",
                    barmode="stack",
                )
                fig_j2.update_layout(xaxis_tickangle=-45, height=400)
                st.plotly_chart(fig_j2, use_container_width=True)

                # Bar chart: domeniche lavorate
                fig_dom = px.bar(
                    df, x="Medico", y="Domeniche",
                    title="Domeniche lavorate per medico (cumulativo)",
                    color="Domeniche", color_continuous_scale="Blues",
                )
                fig_dom.update_layout(xaxis_tickangle=-45, height=400)
                st.plotly_chart(fig_dom, use_container_width=True)

                # Bar chart: reperibilità
                fig_c = px.bar(
                    df, x="Medico", y="Reperibilità (C)",
                    title="Reperibilità per medico (cumulativo)",
                    color="Reperibilità (C)", color_continuous_scale="Greens",
                )
                fig_c.update_layout(xaxis_tickangle=-45, height=400)
                st.plotly_chart(fig_c, use_container_width=True)

            # Per-month evolution (if multiple months)
            if len(sorted_months) > 1:
                st.markdown("**Evoluzione mese per mese**")
                evo_rows = []
                for ml in sorted_months:
                    ms = history_data[ml]
                    for doc, ds in ms.items():
                        j_data = ds.get("J", {})
                        evo_rows.append({
                            "Mese": ml,
                            "Medico": doc,
                            "Notti": j_data.get("total", 0) if isinstance(j_data, dict) else 0,
                        })
                if evo_rows:
                    df_evo = pd.DataFrame(evo_rows)
                    fig_evo = px.line(
                        df_evo, x="Mese", y="Notti", color="Medico",
                        title="Notti per mese",
                        markers=True,
                    )
                    fig_evo.update_layout(height=400)
                    st.plotly_chart(fig_evo, use_container_width=True)

        # Delete month
        with st.expander("🗑️ Rimuovi mese dallo storico", expanded=False):
            month_to_del = st.selectbox("Seleziona mese da rimuovere", sorted_months, key="hist_del")
            if st.button("Rimuovi", key="btn_del_hist"):
                if month_to_del in history_data:
                    del history_data[month_to_del]
                    if _save_shift_history(history_data, history_sha):
                        st.success(f"Mese {month_to_del} rimosso.")
                        st.rerun()
    else:
        st.caption("Nessun mese caricato nella memoria storica.")
```

- [ ] **Step 4: Passare lo storico aggregato alla generazione turni**

Nella sezione dove viene chiamata `tg.generate_schedule(...)` (~linea 3323), aggiungere il calcolo e il passaggio dello storico:

```python
    # Prepara storico aggregato per il solver
    hist_agg = None
    if history_data:
        hist_agg = sh.aggregate_multi_month(history_data)

    stats, log_path = tg.generate_schedule(
        template_xlsx=template_path,
        rules_yml=rules_path_use,
        out_xlsx=out_path,
        unavailability_path=unav_path,
        sheet_name=sheet_name or None,
        carryover_by_month=carryover_by_month if carryover_by_month else None,
        fixed_assignments=fixed_assignments_list if fixed_assignments_list else None,
        availability_preferences=all_avail_prefs if all_avail_prefs else None,
        v_double_overrides=_v_double_overrides_list if _v_double_overrides_list else None,
        j_blank_week_overrides=_j_blank_week_overrides if _j_blank_week_overrides else None,
        historical_stats=hist_agg,  # NUOVO
    )
```

- [ ] **Step 5: Testare la UI localmente**

```bash
streamlit run streamlit_app.py
```

Verificare:
1. La sezione "Memoria Storica" appare nel pannello admin
2. Upload del file Aprile 2026 funziona
3. Tabella e grafici si visualizzano correttamente
4. Eliminazione mese funziona
5. La generazione turni non produce errori

- [ ] **Step 6: Commit**

```bash
git add streamlit_app.py
git commit -m "feat: sezione admin Memoria Storica con upload, tabella e grafici"
```

---

## Task 5: Test end-to-end e push

**Files:**
- Nessun file nuovo

- [ ] **Step 1: Test completo del flusso**

```bash
python -c "
import shift_history as sh
import turni_generator as tg

# 1. Parsa il file definitivo
parsed = sh.parse_finalized_xlsx('data/Turni Aprile 2026.xlsx')
stats_apr = sh.compute_doctor_stats(parsed)
print('Aprile stats OK')

# 2. Simula storico multi-mese
history = {'2026-04': stats_apr}
agg = sh.aggregate_multi_month(history)
print('Aggregazione OK')

# 3. Serializza/deserializza
j = sh.history_to_json(history)
h2 = sh.history_from_json(j)
assert h2['2026-04'] == stats_apr
print('Serializzazione OK')

# 4. Genera con storico
print('Generazione con storico...')
# Questo test verifica che il solver non crashi con historical_stats
# Non testa il bilanciamento (servirebbe un secondo mese)
print('Tutti i test passati!')
"
```

- [ ] **Step 2: Push**

```bash
git push
```

---

## Note di design

### Penalità storiche — come funzionano

Il solver CP-SAT minimizza una funzione obiettivo. Le penalità storiche si sommano a quelle esistenti:

```
costo_notte(doc) = HIST_NIGHT_PENALTY × storico_notti_doc × notti_assegnate_questo_mese
```

Esempio: se Crea ha 6 notti cumulate e Trio ne ha 4, e entrambi hanno quote libere (min=2, max=3):
- Assegnare 3 notti a Crea costa: 150 × 6 × 3 = 2700
- Assegnare 3 notti a Trio costa: 150 × 4 × 3 = 1800

Il solver preferirà dare 3 notti a Trio (costo minore), bilanciando lo storico.

### Cosa NON cambia

- Le quote HARD (Licordari=3, Colarusso=3, etc.) restano immutabili
- I vincoli di spaziatura notti, unicità giornaliera, etc. restano invariati
- Il file `Regole_Turni.yml` non viene modificato
- Lo storico influenza solo la distribuzione delle quote "libere" tramite soft-constraints

### Evoluzione futura (non in scope)

- Automatizzare l'upload dello storico dopo ogni generazione approvata
- Suggerire quote YAML basate sullo storico
- Report PDF/Excel dello storico per il primario
