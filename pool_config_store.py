# -*- coding: utf-8 -*-
"""Pool configuration store — overlay JSON su Regole_Turni.yml.

Gestisce load/save/validate/migrate del file data/pool_config.json su GitHub.
Struttura analoga a shift_history.py (storage JSON) e unavailability_store.py
(funzioni pure parse/validate).

Il JSON sovrascrive pool, quote e flag al momento della generazione turni.
Il YAML rimane invariato come template avanzato.
"""

from __future__ import annotations

import copy
import json
from datetime import datetime, timezone
from typing import Optional

POOL_CONFIG_PATH_DEFAULT = "data/pool_config.json"
SCHEMA_VERSION = 1

QUOTA_TYPES = {"fixed", "max", "min"}
COMBINATION_MODES = {"always", "fallback", "preferred"}
FREE_COLUMNS = {"AD", "AE", "AF", "AG"}
AUTO_COLUMNS = {"AA", "AC"}

_COL_RULE_TARGETS: dict[str, list[tuple[str, str | None]]] = {
    "C": [("C_reperibilita", None)],
    # D/F share one solver rule, but D is the primary ward column.
    # Doctors enabled only on F must not become primary D/F pair doctors.
    "D": [("D_F", "allowed")],
    "F": [],
    "E": [("E_G", "allowed")],
    "G": [("E_G", "allowed")],
    "H": [("H", "pool_mon_fri"), ("H", "distribution_pool")],
    "I": [("I", "distribution_pool")],
    "J": [("J", "pool_other")],
    "K": [("K", "pool")],
    "L": [("L", "pool_other")],
    "Q": [("Q", "pool")],
    "R": [("R", "pool")],
    "S": [("S", "pool")],
    "T": [("T", "pool")],
    "U": [("U", "pool")],
    "V": [("V", "pool")],
    "W": [("W", "other_days_pool")],
    "Y": [("Y", "other_pool")],
    "Z": [("Z", "pool")],
    "AB": [("AB", "fallback_pool")],
}

_DOCTOR_REQUIRED_KEYS = {
    "active",
    "columns",
    "festivi_diurni",
    "festivi_notti",
    "exclude_saturday_day",
    "excluded_from_reperibilita",
    "university_doctor",
    "column_overrides",
}


# ── 1. Serializzazione pura ──────────────────────────────────────────────────

def pool_config_to_text(cfg: dict) -> str:
    return json.dumps(normalize_pool_config(cfg), ensure_ascii=False, indent=2)


def pool_config_from_text(text: str) -> dict:
    return normalize_pool_config(json.loads(text))


def normalize_pool_config(cfg: dict) -> dict:
    """Normalize historical JSON quirks before validation/save."""
    if not isinstance(cfg, dict):
        return cfg
    out = copy.deepcopy(cfg)
    doctors = out.get("doctors")
    if isinstance(doctors, dict):
        for dcfg in doctors.values():
            if not isinstance(dcfg, dict):
                continue
            if isinstance(dcfg.get("columns"), list):
                dcfg["columns"] = sorted({
                    str(c).strip().upper()
                    for c in dcfg.get("columns", [])
                    if str(c).strip()
                    and str(c).strip().upper() not in FREE_COLUMNS
                    and str(c).strip().upper() not in AUTO_COLUMNS
                    and str(c).strip().upper() != "C"
                })
            dcfg["exclude_saturday_day"] = bool(dcfg.get("exclude_saturday_day", False))
            overrides = dcfg.get("column_overrides")
            if isinstance(overrides, dict):
                normalized_overrides = {}
                for col, ov in overrides.items():
                    c = str(col).strip().upper()
                    if c:
                        normalized_overrides[c] = ov
                dcfg["column_overrides"] = normalized_overrides
    combos = out.get("service_combinations")
    if isinstance(combos, list):
        for combo in combos:
            if not isinstance(combo, dict):
                continue
            cols_list = [str(c).strip().upper() for c in (combo.get("columns") or []) if str(c).strip()]
            combo["columns"] = cols_list
            cols = tuple(sorted(cols_list))
            if cols == ("K", "T") and combo.get("mode") == "always":
                combo["mode"] = "fallback"
    critical = out.get("critical_services")
    if isinstance(critical, dict):
        normalized_critical = {}
        for col, spec in critical.items():
            c = str(col).strip().upper()
            if c:
                normalized_critical[c] = spec
        out["critical_services"] = normalized_critical
    col_settings = out.get("column_settings")
    if isinstance(col_settings, dict):
        normalized_settings = {}
        for col, spec in col_settings.items():
            c = str(col).strip().upper()
            if c:
                normalized_settings[c] = spec
        out["column_settings"] = normalized_settings
    return out


# ── 2. Skeleton vuoto ────────────────────────────────────────────────────────

def empty_pool_config() -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "doctors": {},
        "column_settings": {
            "J": {
                "monthly_target": 2,
                "spacing_min_days": 5,
                "spacing_preferred_days": 7,
                "counts_as": 2,
            },
            "C": {
                "monthly_target": None,
                "spacing_min_days": 3,
                "counts_as": 0,
            },
        },
        "service_combinations": [],
        "critical_services": {},
        "updated_at": _now_iso(),
        "updated_by": "admin",
    }


# ── 3. Validazione ───────────────────────────────────────────────────────────

def validate_pool_config(cfg: dict, cfg_yaml: Optional[dict] = None) -> list[str]:
    """Ritorna lista di errori (vuota se OK)."""
    cfg = normalize_pool_config(cfg)
    audit = audit_pool_config(cfg, cfg_yaml)
    return list(audit["errors"])


def audit_pool_config(cfg: dict, cfg_yaml: Optional[dict] = None) -> dict:
    """Validate pool_config and return errors, warnings and derived previews."""
    cfg = normalize_pool_config(cfg)
    errs: list[str] = []
    warnings: list[str] = []
    preview: dict = {"column_pools": {}}

    if not isinstance(cfg, dict):
        return {"errors": ["La configurazione non è un dizionario valido"], "warnings": [], "preview": preview}

    from turni_generator import norm_name

    if cfg.get("schema_version") != SCHEMA_VERSION:
        errs.append(f"schema_version deve essere {SCHEMA_VERSION}, trovato: {cfg.get('schema_version')}")

    yaml_columns = set((cfg_yaml or {}).get("columns", {}).keys())
    editable_columns = (yaml_columns - FREE_COLUMNS - AUTO_COLUMNS) if yaml_columns else set(_COL_RULE_TARGETS.keys())
    known_columns = set(yaml_columns or editable_columns) | set(_COL_RULE_TARGETS.keys())
    rules = (cfg_yaml or {}).get("rules", {}) if isinstance((cfg_yaml or {}).get("rules", {}), dict) else {}
    never_in_j = {
        norm_name(d)
        for d in ((rules.get("J") or {}).get("never_in_J") or ["De Gregorio", "Manganaro"])
    }

    doctors = cfg.get("doctors", {})
    active_docs: dict[str, dict] = {}
    known_doc_norms: dict[str, str] = {}
    if not isinstance(doctors, dict):
        errs.append("'doctors' deve essere un dizionario")
    else:
        for name, dcfg in doctors.items():
            name_s = str(name).strip()
            dn = norm_name(name_s)
            if not name_s:
                errs.append("Nome medico vuoto")
                continue
            if dn in known_doc_norms and known_doc_norms[dn] != name_s:
                errs.append(f"Medico duplicato dopo normalizzazione: '{known_doc_norms[dn]}' e '{name_s}'")
            known_doc_norms[dn] = name_s
            if not isinstance(dcfg, dict):
                errs.append(f"Medico '{name}': deve essere un dizionario")
                continue
            missing = _DOCTOR_REQUIRED_KEYS - dcfg.keys()
            if missing:
                errs.append(f"Medico '{name}': campi mancanti: {sorted(missing)}")
            if dcfg.get("active", True):
                active_docs[name_s] = dcfg
            columns = dcfg.get("columns", [])
            if not isinstance(columns, list):
                errs.append(f"Medico '{name}': 'columns' deve essere una lista")
            else:
                cols_norm = [str(c).strip().upper() for c in columns if str(c).strip()]
                for col in cols_norm:
                    if col in FREE_COLUMNS or col in AUTO_COLUMNS or col == "C":
                        errs.append(f"Medico '{name}': colonna '{col}' non modificabile da pool_config")
                    elif col not in editable_columns:
                        errs.append(f"Medico '{name}': colonna '{col}' sconosciuta o non esposta in GUI")
                if dn in never_in_j and "J" in cols_norm:
                    errs.append(f"Medico '{name}': non puo' essere abilitato in J perche' e' in J.never_in_J")
                if dcfg.get("active", True) and not cols_norm and dcfg.get("excluded_from_reperibilita", False):
                    warnings.append(f"Medico '{name}' attivo ma senza colonne e senza reperibilita'")
            overrides = dcfg.get("column_overrides", {})
            if not isinstance(overrides, dict):
                errs.append(f"Medico '{name}': 'column_overrides' deve essere un dizionario")
            else:
                for col, ov in overrides.items():
                    col_u = str(col).strip().upper()
                    if col_u not in editable_columns:
                        errs.append(f"Medico '{name}', colonna '{col_u}': override su colonna sconosciuta/non esposta")
                    if dn in never_in_j and col_u == "J":
                        errs.append(f"Medico '{name}': non puo' avere override su J perche' e' in J.never_in_J")
                    if not isinstance(ov, dict):
                        errs.append(f"Medico '{name}', colonna '{col}': override deve essere un dizionario")
                        continue
                    qt = ov.get("quota_type")
                    if qt is not None and qt not in QUOTA_TYPES:
                        errs.append(f"Medico '{name}', colonna '{col}': quota_type '{qt}' non valido (ammessi: {sorted(QUOTA_TYPES)})")
                    mq = ov.get("monthly_quota")
                    if mq is not None and (not isinstance(mq, int) or mq < 0):
                        errs.append(f"Medico '{name}', colonna '{col}': monthly_quota deve essere intero >= 0")

    combos = cfg.get("service_combinations", [])
    if not isinstance(combos, list):
        errs.append("'service_combinations' deve essere una lista")
    else:
        for i, combo in enumerate(combos):
            if not isinstance(combo, dict):
                errs.append(f"service_combinations[{i}]: deve essere un dizionario")
                continue
            cols = combo.get("columns", [])
            if not isinstance(cols, list) or len(cols) != 2:
                errs.append(f"service_combinations[{i}]: 'columns' deve essere una lista di 2 lettere")
            else:
                cols_u = [str(c).strip().upper() for c in cols]
                for col in cols_u:
                    if col not in editable_columns:
                        errs.append(f"service_combinations[{i}]: colonna '{col}' sconosciuta/non esposta")
            mode = combo.get("mode")
            if mode not in COMBINATION_MODES:
                errs.append(f"service_combinations[{i}]: mode '{mode}' non valido (ammessi: {sorted(COMBINATION_MODES)})")
            if isinstance(cols, list) and tuple(sorted(str(c).strip().upper() for c in cols)) == ("K", "T") and mode != "fallback":
                warnings.append("K/T viene sempre trattato come fallback di emergenza, non come combinazione abituale")

    critical = cfg.get("critical_services", {})
    if not isinstance(critical, dict):
        errs.append("'critical_services' deve essere un dizionario")
    else:
        for col, spec in critical.items():
            if not isinstance(spec, dict):
                errs.append(f"critical_services['{col}']: deve essere un dizionario")
                continue
            fb = spec.get("fallback")
            if fb != "any" and not isinstance(fb, list):
                errs.append(f"critical_services['{col}']: fallback deve essere 'any' o una lista di medici")
            col_u = str(col).strip().upper()
            if col_u not in editable_columns:
                errs.append(f"critical_services['{col_u}']: colonna sconosciuta/non esposta")
            if isinstance(fb, list):
                for doc in fb:
                    dn = norm_name(doc)
                    if dn not in known_doc_norms:
                        errs.append(f"critical_services['{col_u}']: medico fallback sconosciuto '{doc}'")
                    elif not active_docs.get(known_doc_norms[dn]):
                        errs.append(f"critical_services['{col_u}']: medico fallback non attivo '{doc}'")

    col_settings = cfg.get("column_settings", {})
    if not isinstance(col_settings, dict):
        errs.append("'column_settings' deve essere un dizionario")
    else:
        for col, col_cfg in col_settings.items():
            col_u = str(col).strip().upper()
            if col_u not in editable_columns:
                errs.append(f"column_settings.{col_u}: colonna sconosciuta/non esposta")
            if not isinstance(col_cfg, dict):
                errs.append(f"column_settings.{col_u}: deve essere un dizionario")
                continue
            spacing_min = int(col_cfg.get("spacing_min_days", 0) or 0)
            spacing_pref = int(col_cfg.get("spacing_preferred_days", 0) or 0)
            if spacing_pref and spacing_pref < spacing_min:
                warnings.append(f"column_settings.{col_u}: spacing preferito minore dello spacing minimo")
            if col_cfg.get("monthly_target") not in (None, ""):
                try:
                    if int(col_cfg.get("monthly_target")) < 0:
                        errs.append(f"column_settings.{col_u}.monthly_target deve essere >= 0")
                except Exception:
                    errs.append(f"column_settings.{col_u}.monthly_target deve essere un intero")
        c_cfg = col_settings.get("C", {})
        if isinstance(c_cfg, dict):
            ca = c_cfg.get("counts_as")
            if ca is not None and ca != 0:
                errs.append("column_settings.C.counts_as deve essere 0 (reperibilità non conta nel workload)")

    if isinstance(doctors, dict):
        for col in sorted(editable_columns & set(_COL_RULE_TARGETS.keys())):
            if col == "C":
                continue
            pool = [
                doc for doc, dcfg in doctors.items()
                if isinstance(dcfg, dict)
                and dcfg.get("active", True)
                and col in [str(c).strip().upper() for c in (dcfg.get("columns") or [])]
            ]
            preview["column_pools"][col] = pool
            if col in critical and not pool:
                errs.append(f"Colonna indispensabile {col}: pool primario vuoto")
            elif not pool:
                warnings.append(f"Colonna {col}: pool vuoto; il solver potra' lasciarla vuota o usare solo fallback tecnici")

    return {"errors": errs, "warnings": warnings, "preview": preview}


# ── 4. GitHub storage ────────────────────────────────────────────────────────

def load_pool_config_from_github(
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    path: str = POOL_CONFIG_PATH_DEFAULT,
) -> tuple[dict, Optional[str]]:
    """Carica pool_config da GitHub. Ritorna ({}, None) se il file non esiste."""
    from github_utils import get_file

    gf = get_file(owner, repo, path, token, branch)
    if gf is None:
        return {}, None
    return pool_config_from_text(gf.text), gf.sha


def save_pool_config_to_github(
    cfg: dict,
    owner: str,
    repo: str,
    token: str,
    branch: str = "main",
    sha: Optional[str] = None,
    path: str = POOL_CONFIG_PATH_DEFAULT,
) -> dict:
    """Salva pool_config su GitHub. Ritorna risposta GitHub Contents API."""
    from github_utils import put_file

    text = pool_config_to_text(cfg)
    return put_file(
        owner,
        repo,
        path,
        token,
        "Aggiornamento configurazione pool medici",
        text,
        branch,
        sha,
    )


# ── 5. Migrazione da YAML ────────────────────────────────────────────────────

def migrate_from_yaml(cfg_yaml: dict) -> dict:
    """Costruisce un pool_config iniziale leggendo il YAML corrente.

    L'admin rifinisce i dettagli dopo la migrazione — questo è un punto di
    partenza automatico che rispecchia la configurazione attuale.
    """
    from turni_generator import collect_doctors, norm_name

    rules = cfg_yaml.get("rules", {})
    gc = cfg_yaml.get("global_constraints", {})
    abs_excl = {norm_name(x) for x in (cfg_yaml.get("absolute_exclusions") or [])}

    rJ = rules.get("J", {})
    rC = rules.get("C_reperibilita", {})
    rFest = rules.get("Festivi", {})

    j_pool_other = {norm_name(x) for x in (rJ.get("pool_other") or [])}
    j_monthly_quotas = rJ.get("monthly_quotas") or {}
    j_weekend_excluded = {norm_name(x) for x in (rJ.get("weekend_excluded_doctors") or [])}

    c_excluded = {norm_name(x) for x in (rC.get("excluded") or [])}

    fest_excl = {norm_name(x) for x in (rFest.get("excluded") or [])}
    fest_pool = {norm_name(x) for x in (rFest.get("pool") or [])}

    gc_uni = gc.get("university_doctors") or {}
    uni_ratio = float(gc.get("university_ratio", 0.6))

    all_doctors = collect_doctors(cfg_yaml)

    # Mappa colonna → chiave pool nel YAML e lista medici
    col_to_doctors = _build_col_to_doctors(rules, all_doctors)

    doctors_cfg: dict = {}
    for doc in all_doctors:
        dn = norm_name(doc)
        active = dn not in abs_excl

        columns = sorted(
            col for col, pool in col_to_doctors.items() if dn in pool
        )

        festivi_diurni = dn not in fest_excl and (dn in fest_pool or active)
        festivi_notti = dn in j_pool_other and dn not in fest_excl

        excluded_from_rep = dn in c_excluded

        uni_cfg = gc_uni.get(doc) or gc_uni.get(dn)
        university_doctor = {"ratio": uni_ratio} if uni_cfg else None

        column_overrides: dict = {}
        mq = j_monthly_quotas.get(doc) or j_monthly_quotas.get(dn)
        if mq is not None:
            column_overrides["J"] = {"monthly_quota": int(mq), "quota_type": "fixed"}
        if dn in j_weekend_excluded:
            column_overrides.setdefault("J", {})["weekend_nights"] = False

        doctors_cfg[doc] = {
            "active": active,
            "columns": columns,
            "festivi_diurni": festivi_diurni,
            "festivi_notti": festivi_notti,
            "exclude_saturday_day": False,
            "excluded_from_reperibilita": excluded_from_rep,
            "university_doctor": university_doctor,
            "column_overrides": column_overrides,
        }

    # Aggiungi medici in absolute_exclusions come active=false (non in collect_doctors)
    for doc in (cfg_yaml.get("absolute_exclusions") or []):
        if doc not in doctors_cfg:
            doctors_cfg[doc] = {
                "active": False,
                "columns": [],
                "festivi_diurni": False,
                "festivi_notti": False,
                "exclude_saturday_day": False,
                "excluded_from_reperibilita": True,
                "university_doctor": None,
                "column_overrides": {},
            }

    relief = gc.get("relief_valves") or {}
    service_combinations = []
    if relief.get("enable_kt_share", False):
        service_combinations.append({"columns": ["K", "T"], "same_day": True, "mode": "fallback"})

    cfg = empty_pool_config()
    cfg["doctors"] = doctors_cfg
    cfg["service_combinations"] = service_combinations
    cfg["updated_at"] = _now_iso()
    return cfg


# ── Helpers interni ──────────────────────────────────────────────────────────

def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _build_col_to_doctors(rules: dict, all_doctors: list[str]) -> dict[str, set[str]]:
    """Costruisce mappa colonna → set di nomi normalizzati dal YAML."""
    from turni_generator import norm_name

    POOL_KEYS = [
        "allowed", "pool", "pool_other", "pool_mon_fri", "other_pool",
        "fallback_pool", "distribution_pool", "other_days_pool",
    ]
    FIXED_KEYS = ["fixed", "tuesday_fixed", "friday_required_doctor"]

    col_to_docs: dict[str, set[str]] = {}

    col_rule_map = {
        "C": "C_reperibilita",
        "D": "D_F", "F": "D_F",
        "E": "E_G", "G": "E_G",
        "H": "H", "I": "I",
        "J": "J",
        "K": "K", "L": "L", "Q": "Q", "R": "R", "S": "S",
        "T": "T", "U": "U", "V": "V", "W": "W",
        "Y": "Y", "Z": "Z", "AA": "AA", "AB": "AB", "AC": "AC",
    }

    dn_all = {norm_name(d) for d in all_doctors}

    for col, rule_key in col_rule_map.items():
        rule = rules.get(rule_key, {})
        if not isinstance(rule, dict):
            continue
        pool: set[str] = set()
        for k in POOL_KEYS:
            if k in rule and isinstance(rule[k], list):
                pool |= {norm_name(x) for x in rule[k] if x}
        for k in FIXED_KEYS:
            if rule.get(k):
                pool.add(norm_name(rule[k]))
        # Per J: includi anche chi ha monthly_quotas (entra nel pool via quota)
        if col == "J":
            mq = rule.get("monthly_quotas") or {}
            pool |= {norm_name(d) for d in mq if d}
        # Includi solo medici riconosciuti da collect_doctors
        col_to_docs[col] = pool & dn_all

    return col_to_docs
