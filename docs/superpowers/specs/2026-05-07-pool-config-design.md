# Pool Config — Spec di Design

**Data:** 2026-05-07 (rev. 2026-05-07b)
**Stato:** Approvato per implementazione

---

## Obiettivo

Permettere all'admin di gestire pool medici, quote, combinazioni di servizi e servizi critici direttamente dalla GUI Streamlit, senza toccare `Regole_Turni.yml` a mano.

---

## Architettura: Overlay JSON su YAML

`Regole_Turni.yml` rimane invariato come template avanzato (spacing, penalità, relief_valves). La GUI legge e scrive un file separato `data/pool_config.json` su GitHub. Al momento della generazione turni, `apply_pool_config(cfg_yaml, pool_config)` in `turni_generator.py` fa il merge: il JSON sovrascrive pool, quote e flag; il YAML fornisce tutti gli altri parametri del solver.

**Vantaggio:** rollback immediato (basta non usare il JSON), nessuna rottura del YAML, i vincoli avanzati restano nel YAML modificabile a mano per casi eccezionali.

---

## Cosa NON è gestito da pool_config (già gestito altrove)

| Funzionalità | Dove vive | Note |
|---|---|---|
| Giorno vuoto Notte (J) per settimana | GUI esistente riga ~3716 (`j_blank_week_overrides`) | `thursday_blank` rimane nel YAML come flag di attivazione |
| Turno doppio Sala PM (V) per settimana | GUI esistente riga ~3637 (`v_double_overrides`) | Venerdì Crea + override per settimana già funzionante |
| Vincoli strutturali fissi (vedi sezione dedicata) | YAML | Panel sola lettura in GUI |

---

## Schema `pool_config.json`

```json
{
  "schema_version": 1,
  "doctors": {
    "Licordari": {
      "active": true,
      "columns": ["C", "D", "E", "J", "K", "T", "Q", "L", "R", "Y"],
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
      "columns": ["D", "E", "J", "H", "I", "K", "T", "Q", "L", "R"],
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
      "columns": ["J", "Q", "R", "S"],
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
      "columns": ["J", "E", "U", "V", "AB"],
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
      "columns": ["D", "E", "H", "I", "K", "Z"],
      "festivi_diurni": true,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
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
    },
    "Grimaldi": {
      "active": true,
      "columns": ["D", "F"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": null,
      "column_overrides": {}
    },
    "Recupero": {
      "active": true,
      "columns": ["Q", "R", "T", "L", "W", "Y"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": null,
      "column_overrides": {
        "T": { "monthly_quota": 4, "quota_type": "min" },
        "Q": { "monthly_quota": 2, "quota_type": "max" },
        "W": { "monthly_quota": 8, "quota_type": "max" }
      }
    }
  },
  "column_settings": {
    "J": {
      "monthly_target": 2,
      "spacing_min_days": 5,
      "spacing_preferred_days": 7,
      "counts_as": 2
    },
    "C": {
      "monthly_target": null,
      "spacing_min_days": 3,
      "counts_as": 0
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
    "H": { "fallback": ["Licordari", "Allegra", "Cimino"] }
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
| `columns` | list[str] | Colonne Excel in cui può essere assegnato (pool primario). Se J non è in lista, il medico non fa notti. |
| `festivi_diurni` | bool | Partecipa al pool festivi D/E/H/I |
| `festivi_notti` | bool | Partecipa al pool notti J nei giorni festivi |
| `excluded_from_reperibilita` | bool | Non viene mai assegnato a C (reperibilità) |
| `university_doctor` | obj\|null | Se presente: `{ "ratio": 0.6 }` — il solver applica il ratio al calcolo del workload mensile (Mon-Sat, esclusi festivi) |
| `column_overrides` | dict | Override per colonna specifica (vedi sotto) |

### `column_overrides[col_letter]`

| Campo | Tipo | Significato |
|---|---|---|
| `monthly_quota` | int | Numero di turni mensili per quella colonna |
| `quota_type` | `"fixed"` \| `"max"` \| `"min"` | `fixed` = sempre esattamente N; `max` = mai più di N; `min` = almeno N |
| `weekend_nights` | bool | Solo per J: se false, escluso da notti di sabato e domenica |

### `column_settings[col_letter]`

| Campo | Tipo | Significato |
|---|---|---|
| `monthly_target` | int\|null | Quota mensile di default per tutti i medici senza override. `null` = distribuzione automatica equa |
| `spacing_min_days` | int | Giorni minimi (vincolo hard) tra due assegnazioni consecutive sulla stessa colonna |
| `spacing_preferred_days` | int | Giorni preferiti (vincolo soft) tra due assegnazioni — solo per J (valore: 7) |
| `counts_as` | int | Peso del turno nel conteggio totale. J=2 globale. C=0 (reperibilità non conta nel workload) |

**Nota C:** `counts_as` per C è bloccato a 0 nella GUI (non editabile). La reperibilità non rientra nel conteggio turni.

### `service_combinations[]`

| Campo | Tipo | Significato |
|---|---|---|
| `columns` | list[str] | Colonne che condividono il medico nello stesso giorno |
| `same_day` | bool | Sempre true in questa versione |
| `mode` | str | `"always"` = vincolo hard; `"fallback"` = solo se pool esaurito; `"preferred"` = soft penalty |

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
| Calabrò | `weekend_nights: false` | Fa notti feriali, escluso sab/dom |
| Grimaldi | J non in `columns` | Non fa notti |
| Tutti gli altri | nessuno | Target flessibile 2; possono fare 3 a rotazione se necessario |

---

## Vincoli strutturali fissi (rimangono nel YAML)

Questi vincoli non cambiano mese per mese — dipendono dalla struttura del reparto. **Non sono modificabili dalla GUI** ma vengono mostrati in un pannello informativo sola lettura nella sezione Gestione Pool.

| Vincolo | YAML key | Descrizione |
|---|---|---|
| Cimino esatto 2 U al mese | `U.cimino_exact_per_month: 2` | Vincolo hard su colonna U |
| Crea unico sabato AB (2/mese) | `AB.saturday_only_doctor: Crea` | Solo Crea nei sabati AB |
| Allegra lunedì V+U | `U.v_allegra_monday_constraint: true` | Vincolo same-day lunedì |
| De Gregorio max 3 feriali su I | `I.degregorio_max_weekdays: 3` | Cap per-medico su colonna I |
| Grimaldi/Calabrò esenti min weekend liberi | `weekend_off_exempt: [Grimaldi, Calabrò]` | Non soggetti al vincolo "min 2 weekend liberi/mese" |

---

## GUI — Layout admin (4 tab, PIN-protetti)

### Tab 1 — Medici
Tabella con tutti i medici: attivo, reperibilità, festivi diurni, festivi notti, universitario, colonne assegnate. Click su riga apre pannello dettaglio con tutti i campi editabili incluso ratio universitario.

In fondo al tab: pannello sola lettura **"Vincoli strutturali fissi"** — elenca i vincoli della tabella sopra con spiegazione testuale. L'admin sa che esistono ma non può modificarli da qui.

### Tab 2 — Colonne
Selettore medico + griglia di tutti i servizi (da lista fissa derivata dal YAML). Click su un servizio lo aggiunge/rimuove dal pool del medico (toggle visivo). Per C (Reperibilità): `counts_as` mostrato ma bloccato a 0.

### Tab 3 — Limiti
Impostazioni globali per colonna (`monthly_target`, `spacing_min_days`, `spacing_preferred_days`, `counts_as`) + tabella override per singolo medico (`monthly_quota`, `quota_type`: fixed/max/min, `weekend_nights` per J).

### Tab 4 — Servizi
Combinazioni same-day con modalità (always/fallback/preferred) + servizi critici con fallback configurabile (any o lista esplicita).

**Salvataggio:** unico bottone "Salva configurazione" in fondo alla pagina — scrive l'intero `pool_config.json` su GitHub in un'unica operazione atomica.

---

## Integrazione Solver

### Funzione `apply_pool_config(cfg_yaml, pool_config) -> cfg_merged`

1. **Pool per colonna**: per ogni colonna, sostituisce il pool con i medici `active=true` che hanno quella colonna in `columns`. Se nessun medico ha quella colonna, lascia il pool del YAML invariato.
2. **Pool festivi**: filtra i medici con `festivi_diurni=true` per D/E/H/I festivi; `festivi_notti=true` per J nei festivi.
3. **Reperibilità (C)**: esclude i medici con `excluded_from_reperibilita=true` dalla lista `C_reperibilita.excluded`.
4. **Weekend J**: aggiunge i medici con `weekend_nights=false` in `column_overrides.J` alla lista `weekend_excluded_doctors`.
5. **Quote J**:
   - `quota_type=fixed`: vincolo hard CP-SAT `sum(x_J_doc) == N`
   - `quota_type=max`: vincolo hard CP-SAT `sum(x_J_doc) <= N`
   - `quota_type=min`: vincolo hard CP-SAT `sum(x_col_doc) >= N`
   - Senza override: il solver bilancia puntando al `monthly_target`, con slack +1 a rotazione per il gruppo flessibile
6. **`counts_as`**: sostituisce `night_counts_double`. J=2 globale. C=0 (non conta nel workload totale).
7. **`spacing_preferred_days`**: aggiunto come soft constraint per J (penalty per spaziatura < 7 giorni, dopo il hard min di 5).
8. **Combinazioni**:
   - `mode=always`: vincolo hard che lega le due variabili CP-SAT per lo stesso giorno
   - `mode=fallback`: aggiunge la combinazione a `relief_valves` con alta penalità
   - `mode=preferred`: aggiunge soft penalty (500 punti) per medici diversi sulle due colonne
9. **Servizi critici**: aggiunge `emergency_pool` per colonna con penalità molto alta, attivato solo se il primario è esaurito.

### Nuovo file `pool_config_store.py`

```python
def load_pool_config(gf_text: str) -> dict: ...
def save_pool_config(config: dict) -> str: ...          # returns JSON string
def validate_pool_config(config: dict) -> list[str]: ... # returns error list
def migrate_from_yaml(cfg_yaml: dict) -> dict: ...       # popola pool_config da YAML esistente
```

---

## Storage GitHub

- Path: `data/pool_config.json` (configurabile via secrets: `github_unavailability.pool_config_path`)
- Stesso meccanismo get_file/put_file già usato per unavailability e availability
- SHA-based conflict detection
- Un solo file JSON, scrive l'admin, raramente

---

## Migrazione da YAML

Al primo accesso, se `pool_config.json` non esiste, un pulsante "Inizializza da YAML attuale" chiama `migrate_from_yaml(cfg_yaml)` che popola automaticamente i pool leggendo le colonne del YAML esistente. L'admin poi aggiusta i dettagli.

---

## File coinvolti

| File | Modifica |
|---|---|
| `pool_config_store.py` | NUOVO — load/save/validate/migrate |
| `turni_generator.py` | `apply_pool_config()` + generalizzazione `df_pair` → service_combinations + logica critical_services + quota fixed/max/min + spacing_preferred |
| `streamlit_app.py` | Nuova sezione admin "Gestione Pool" (4 tab) + load/save pool_config + bottone inizializza da YAML |
| `github_utils.py` | Nessuna modifica |
| `Regole_Turni.yml` | Non modificato — rimane template avanzato |

---

## Vincoli e assunzioni

- La lista dei servizi nel Tab 2 è derivata dal YAML al momento del caricamento. L'admin non può creare nuove colonne dalla GUI.
- Se un medico ha una colonna in `pool_config` ma non nel YAML, la colonna viene ignorata nel merge.
- `apply_pool_config` è idempotente.
- `thursday_blank` nel YAML e la GUI esistente per `j_blank_week_overrides` (riga ~3716) rimangono invariati — non interferiscono con pool_config.
- La GUI esistente per `v_double_overrides` (riga ~3637) rimane invariata — non interferisce con pool_config.
