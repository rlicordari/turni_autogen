# Pool Config — Spec di Design

**Data:** 2026-05-07  
**Stato:** Approvato per implementazione

---

## Obiettivo

Permettere all'admin di gestire pool medici, quote, combinazioni di servizi e servizi critici direttamente dalla GUI Streamlit, senza toccare `Regole_Turni.yml` a mano.

---

## Architettura: Overlay JSON su YAML

`Regole_Turni.yml` rimane invariato come template avanzato (spacing, penalità, relief_valves). La GUI legge e scrive un file separato `data/pool_config.json` su GitHub. Al momento della generazione turni, `apply_pool_config(cfg_yaml, pool_config)` in `turni_generator.py` fa il merge: il JSON sovrascrive pool, quote e flag; il YAML fornisce tutti gli altri parametri del solver.

**Vantaggio:** rollback immediato (basta non usare il JSON), nessuna rottura del YAML, i vincoli avanzati restano nel YAML modificabile a mano per casi eccezionali.

---

## Schema `pool_config.json`

```json
{
  "schema_version": 1,
  "doctors": {
    "Licordari": {
      "active": true,
      "columns": ["C", "D", "E", "J", "K", "T", "Q"],
      "festivi_diurni": true,
      "festivi_notti": true,
      "excluded_from_reperibilita": false,
      "university_doctor": null,
      "column_overrides": {
        "J": { "monthly_quota": 3, "quota_type": "fixed" }
      }
    },
    "Colarusso": {
      "active": true,
      "columns": ["D", "E", "J", "H", "I"],
      "festivi_diurni": true,
      "festivi_notti": true,
      "excluded_from_reperibilita": false,
      "university_doctor": null,
      "column_overrides": {
        "J": { "monthly_quota": 3, "quota_type": "fixed" }
      }
    },
    "Zito": {
      "active": true,
      "columns": ["J"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": { "ratio": 0.6 },
      "column_overrides": {
        "J": { "monthly_quota": 2, "quota_type": "max" }
      }
    },
    "Dattilo": {
      "active": true,
      "columns": ["J"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": { "ratio": 0.6 },
      "column_overrides": {
        "J": { "monthly_quota": 2, "quota_type": "max" }
      }
    },
    "De Gregorio": {
      "active": true,
      "columns": ["D", "E", "H", "I"],
      "festivi_diurni": true,
      "festivi_notti": false,
      "excluded_from_reperibilita": false,
      "university_doctor": { "ratio": 0.6 },
      "column_overrides": {}
    },
    "Calabrò": {
      "active": true,
      "columns": ["D", "F", "J"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": null,
      "column_overrides": {
        "J": { "weekend_nights": false }
      }
    }
  },
  "column_settings": {
    "J": {
      "monthly_target": 2,
      "spacing_min_days": 5,
      "counts_as": 2
    }
  },
  "service_combinations": [
    { "columns": ["K", "T"], "same_day": true, "mode": "always" },
    { "columns": ["Q", "R"], "same_day": true, "mode": "fallback" }
  ],
  "critical_services": {
    "J": { "fallback": "any" },
    "D": { "fallback": "any" },
    "E": { "fallback": "any" },
    "H": { "fallback": "any" }
  },
  "updated_at": "2026-05-07T10:00:00Z",
  "updated_by": "admin"
}
```

---

## Campi — Dettaglio

### `doctors[name]`

| Campo | Tipo | Significato |
|---|---|---|
| `active` | bool | Se false, il medico non viene mai assegnato (equivale a `absolute_exclusions`) |
| `columns` | list[str] | Colonne Excel in cui può essere assegnato (pool primario) |
| `festivi_diurni` | bool | Partecipa al pool festivi D/E/H/I |
| `festivi_notti` | bool | Partecipa al pool notti J nei giorni festivi |
| `excluded_from_reperibilita` | bool | Non viene mai assegnato a C (reperibilità) |
| `university_doctor` | obj\|null | Se presente: `{ "ratio": 0.6 }` — il solver applica il ratio al calcolo del workload mensile (Mon-Sat, esclusi festivi). Il ratio è già implementato nel solver. |
| `column_overrides` | dict | Override per colonna specifica (vedi sotto) |

### `column_overrides[col_letter]`

| Campo | Tipo | Significato |
|---|---|---|
| `monthly_quota` | int | Numero di turni mensili per quella colonna |
| `quota_type` | `"fixed"` \| `"max"` | `fixed` = sempre esattamente N; `max` = mai più di N (può fare meno se indisponibile) |
| `weekend_nights` | bool | Solo per J: se false, escluso da notti di sabato e domenica |

### `column_settings[col_letter]`

| Campo | Tipo | Significato |
|---|---|---|
| `monthly_target` | int\|null | Quota mensile di default per tutti i medici senza override. `null` = distribuzione automatica equa |
| `spacing_min_days` | int | Giorni minimi tra due assegnazioni consecutive sulla stessa colonna |
| `counts_as` | int | Peso del turno nel conteggio totale (J=2 per tutti, altri=1) |

### `service_combinations[]`

| Campo | Tipo | Significato |
|---|---|---|
| `columns` | list[str] | Colonne che possono condividere il medico nello stesso giorno |
| `same_day` | bool | Sempre true in questa versione |
| `mode` | str | `"always"` = vincolo hard (stesso medico obbligatorio ogni giorno); `"fallback"` = solo se pool esaurito (equivale all'attuale `relief_valves`); `"preferred"` = soft penalty per tenerli separati |

### `critical_services{col: {fallback}}`

| `fallback` | Comportamento |
|---|---|
| `"any"` | Se pool primario esaurito, il solver usa qualsiasi medico attivo |
| `["Doc1","Doc2"]` | Se pool primario esaurito, prova solo questi medici come backup |

---

## Logica quota notti J — Riepilogo

| Medico | Override | Comportamento |
|---|---|---|
| Licordari, Colarusso | `quota: 3, type: fixed` | Sempre esattamente 3 notti — vincolo hard |
| Zito, Dattilo | `quota: 2, type: max` | Mai più di 2 notti — vincolo hard; ratio universitario già applicato dal solver |
| De Gregorio | nessuno + J non in `columns` | Escluso dalle notti (non è nel pool J) |
| Calabrò | `weekend_nights: false` | Fa notti nei feriali, escluso sab/dom |
| Tutti gli altri | nessuno | Target flessibile 2 (da `column_settings.J.monthly_target`); possono fare 3 a rotazione se necessario per coprire il mese |

---

## GUI — Layout admin (4 tab, PIN-protetti)

### Tab 1 — Medici
Tabella con tutti i medici: attivo, reperibilità, festivi diurni, festivi notti, universitario, colonne assegnate. Click su riga apre pannello dettaglio con tutti i campi editabili incluso ratio universitario.

### Tab 2 — Colonne
Selettore medico + griglia di tutti i servizi esistenti (da lista fissa derivata dal YAML). Click su un servizio lo aggiunge/rimuove dal pool del medico (toggle visivo).

### Tab 3 — Limiti
Impostazioni globali per colonna (monthly_target, spacing, counts_as) + tabella override per singolo medico (quota + tipo fisso/massimo + weekend_nights per J). Override disponibile per qualsiasi medico, non solo universitari.

### Tab 4 — Servizi
Combinazioni same-day con modalità (always/fallback/preferred) + servizi critici con fallback configurabile (any o lista esplicita).

**Salvataggio:** unico bottone "Salva configurazione" in fondo alla pagina — scrive l'intero `pool_config.json` su GitHub in un'unica operazione atomica.

---

## Integrazione Solver

### Funzione `apply_pool_config(cfg_yaml, pool_config) -> cfg_merged`

1. **Pool per colonna**: per ogni colonna, sostituisce il pool con i medici `active=true` che hanno quella colonna in `columns`. Se nessun medico ha quella colonna, lascia il pool del YAML invariato.
2. **Pool festivi**: filtra i medici con `festivi_diurni=true` per il pool D/E/H/I festivi; `festivi_notti=true` per J nei festivi.
3. **Reperibilità (C)**: esclude i medici con `excluded_from_reperibilita=true` dalla lista `C_reperibilita.excluded`.
4. **Weekend J**: aggiunge i medici con `weekend_nights=false` alla lista `weekend_excluded_doctors` della colonna J.
5. **Quote J**:
   - Medici con `quota_type=fixed`: vincolo hard CP-SAT `sum(x_J_doc) == N` nel mese.
   - Medici con `quota_type=max`: vincolo hard CP-SAT `sum(x_J_doc) <= N`.
   - Medici senza override: il solver bilancia puntando al `monthly_target`, con slack +1 a rotazione per chi è nel gruppo flessibile.
6. **`counts_as`**: sostituisce `night_counts_double` per i medici universitari. Il valore si applica a tutti, non solo agli universitari (J=2 è globale).
7. **Combinazioni**:
   - `mode=always`: vincolo hard che lega le due variabili CP-SAT per lo stesso giorno (generalizza il meccanismo `df_pair`).
   - `mode=fallback`: aggiunge la combinazione a `relief_valves` con alta penalità (comportamento attuale).
   - `mode=preferred`: aggiunge soft penalty (500 punti) per medici diversi sulle due colonne.
8. **Servizi critici**: aggiunge `emergency_pool` per colonna — pool secondario con penalità molto alta attivato solo se il primario è esaurito. Con `fallback="any"`: tutti i medici attivi. Con lista esplicita: solo quelli nominati.

### Nuovo file `pool_config_store.py`
Funzioni pure per load/save/validate di `pool_config.json`. Analogia con `unavailability_store.py`.

```python
def load_pool_config(gf_text: str) -> dict: ...
def save_pool_config(config: dict) -> str: ...  # returns JSON string
def validate_pool_config(config: dict) -> list[str]: ...  # returns error list
def migrate_from_yaml(cfg_yaml: dict) -> dict: ...  # popola pool_config da YAML esistente
```

---

## Storage GitHub

- Path: `data/pool_config.json` (configurabile via secrets: `github_unavailability.pool_config_path`)
- Stesso meccanismo get_file/put_file già usato per unavailability e availability
- SHA-based conflict detection (stesso admin non può salvare da due schede)
- Nessun file per-medico: un solo file JSON, scrive l'admin, raramente

---

## Migrazione da YAML

Al primo accesso alla sezione Gestione Pool, se `pool_config.json` non esiste, un pulsante "Inizializza da YAML attuale" chiama `migrate_from_yaml(cfg_yaml)` che popola automaticamente i pool leggendo le colonne del YAML esistente. L'admin poi aggiusta i dettagli.

---

## File coinvolti

| File | Modifica |
|---|---|
| `pool_config_store.py` | NUOVO — load/save/validate/migrate |
| `turni_generator.py` | `apply_pool_config()` + generalizzazione `df_pair` → service_combinations + logica critical_services + quota fixed/max |
| `streamlit_app.py` | Nuova sezione admin "Gestione Pool" (4 tab) + load/save pool_config + bottone inizializza da YAML |
| `github_utils.py` | Nessuna modifica (usa get_file/put_file esistenti) |
| `Regole_Turni.yml` | Non modificato — rimane template avanzato |

---

## Vincoli e assunzioni

- La lista dei servizi mostrati nel Tab 2 è derivata dal YAML al momento del caricamento (colonne esistenti). L'admin non può creare nuove colonne dalla GUI.
- Se un medico ha una colonna in `pool_config` ma non nel YAML, la colonna viene ignorata nel merge (il YAML è autoritativo sulla struttura del solver).
- `apply_pool_config` è idempotente: applicato più volte al medesimo YAML + JSON produce sempre lo stesso risultato.
- Il pool_config non gestisce spacing o penalità (rimangono nel YAML) — solo pool, quote e flag.
