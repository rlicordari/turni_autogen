# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Descrizione del progetto

Generatore automatico di turni ospedalieri per UTIC/Cardiologia. Legge un template Excel, un file di regole YAML e un file di indisponibilità mensili; risolve l'assegnazione tramite **OR-Tools CP-SAT** (con fallback greedy); produce un file Excel compilato. Espone anche una **web UI Streamlit** per l'inserimento autonomo delle indisponibilità da parte dei medici e la generazione dei turni da parte dell'admin.

## Installazione

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

Richiede Python 3.10+ (consigliato 3.11).

## Avvio

**CLI:**
```bash
python turni_generator.py \
  --template Turni_Febbraio_2026.xlsx \
  --rules Regole_Turni.yml \
  --unavailability unavailability.xlsx \
  --out Turni_Febbraio_2026_COMPILATI.xlsx
```

**App Streamlit:**
```bash
streamlit run streamlit_app.py
```

**GUI legacy (tkinter):**
```bash
python turni_generator.py --gui
```

## Architettura

| File | Ruolo |
|---|---|
| `turni_generator.py` | Solver principale: legge template Excel + regole YAML + indisponibilità, costruisce il modello CP-SAT, scrive l'Excel di output |
| `streamlit_app.py` | UI Streamlit: login medico (PIN + OTP email), inserimento indisponibilità, generazione turni per admin |
| `unavailability_store.py` | Funzioni pure per il datastore CSV: parsing, filtraggio, deduplicazione, serializzazione |
| `github_utils.py` | Lettura/scrittura via GitHub Contents API (archivia il CSV delle indisponibilità su una repo privata) |
| `xlsx_utils.py` | Genera il file XLSX delle indisponibilità dal CSV usando `unavailability_template.xlsx` |
| `Regole_Turni.yml` | File regole mensile: definizione colonne, pool medici, quote, vincoli, penalità |
| `data/doctor_contacts.yml` | Mappa nome medico → email (usata per gli OTP) |
| `Style_Template.xlsx` | Template di stile opzionale applicato alla generazione di nuovi template Excel mensili |

### Flusso dei dati

1. Le **regole** sono definite in `Regole_Turni.yml` — le lettere di colonna corrispondono ai tipi di turno, ciascuno con pool, quote, vincoli di spaziatura e pesi di penalità.
2. Le **indisponibilità** sono archiviate come CSV in una repo GitHub privata (`data/unavailability_store.csv`). I medici le inseriscono via Streamlit; l'admin può anche caricare un file Excel.
3. Il **solver** (`turni_generator.py`) traduce regole + indisponibilità in variabili e vincoli CP-SAT, risolve e compila il workbook openpyxl.
4. **Streamlit** (`streamlit_app.py`) orchestra il tutto per gli utenti web: gestisce auth PIN, OTP via SMTP, lease di sessione per medico (kick-out in caso di login concorrente) e invoca `turni_generator` in-process.

### Secrets Streamlit (necessari per il funzionamento completo)

```toml
[auth]
admin_pin = "..."

[doctor_pins]
"Cognome" = "1234"

[github_unavailability]
token  = "ghp_..."
owner  = "REPO_OWNER"
repo   = "REPO_NAME"
branch = "main"
path   = "data/unavailability_store.csv"

[smtp]
host     = "smtp.gmail.com"
port     = 587
username = "..."
password = "APP_PASSWORD"
from     = "..."
starttls = true
```

### Keep-alive (GitHub Actions)

`.github/workflows/keep_awake_selenium.yml` fa un ping all'app Streamlit ogni 3 ore tramite Selenium per evitare l'hibernation. Richiede il secret Actions `STREAMLIT_APP_URL`.

## Convenzioni importanti

- I nomi dei medici in `Regole_Turni.yml` devono corrispondere **esattamente** ai nomi nei file di indisponibilità e in `doctor_contacts.yml`.
- Valori ammessi per `Fascia`: `Mattina`, `Pomeriggio`, `Notte`, `Diurno` (= Mattina + Pomeriggio), `Tutto il giorno`.
- Le lettere di colonna (C, D, E, …) nel YAML corrispondono direttamente alle colonne Excel del template.
- `absolute_exclusions` nel YAML elenca i medici mai assegnati ad alcun turno. Attualmente esclusi: De Luca, Carciotto, Virga, Andò, Saporito, D'Angelo.
- **D'Angelo** è stata esclusa temporaneamente (aprile 2026). Per reinserirla, aggiungere "D'Angelo" nei seguenti pool/liste in `Regole_Turni.yml`:
  - `E_G.allowed` (Cardiologia mattina / Riabilitazione)
  - `Q.pool` (ECO base)
  - `T.pool` (Interni)
  - `U.pool` (Contr.PM)
  - `Y.other_pool` (Ambulatori specialistici)
  - `Z.pool` (Vascolare)
  - `AB.fallback_pool` (Holter/Brugada/FA)
  - Rimuoverla da `absolute_exclusions` e da `C_reperibilita.excluded` (se applicabile al mese).
- I medici universitari (`university_doctors`) possono avere `night_counts_double: true` per dimezzare la quota effettiva di notti.
- La sezione `relief_valves` definisce fallback ad alta penalità per evitare l'infeasibility (es. permettere una colonna vuota a costo elevato anziché fallire).

## Feature in sviluppo: Memoria Storica Turni

**Piano completo:** `docs/PLAN_historical_shifts.md`

**Stato avanzamento (aggiornare ad ogni step):**
- [x] Task 1: `shift_history.py` — parser Excel definitivo + aggregazione stats + normalizzazione nomi (commit 2303447)
- [x] Task 2: Storage su GitHub — load/save storico JSON (commit a01a611)
- [x] Task 3: Integrazione solver — soft constraints con `historical_stats`
- [x] Task 4: UI admin Streamlit — upload, tabella, grafici Plotly, eliminazione mese
- [x] Task 5: Test end-to-end e push

**Modifiche sessione 22 aprile 2026:**
- Parser dinamico colonne: `_map_columns_from_header()` legge riga 1 del foglio Excel e mappa header → tag logico (non più posizioni fisse). Supporta layout diversi tra mesi.
- `_HEADER_TO_TAG` in `shift_history.py`: ordine importante — pattern specifici (es. "emodinamica notte") prima di generici ("notte").
- Filtro medici validi: `compute_doctor_stats(parsed, valid_doctors=set)` accetta whitelist da pool YAML per escludere nomi spuri (note, testo libero nelle celle).
- `_EXCLUDED_NAMES` in `shift_history.py`: nomi esclusi a priori dal conteggio (Recupero, De Luca, Saporito, Virga, Carciotto, Andò, D'Angelo).
- `_WEEKEND_COLUMNS = {"C", "D", "E", "H", "I", "J"}` — solo queste colonne per conteggio domeniche/festivi.
- Pasquetta calcolata con `_easter_monday()` (algoritmo Meeus/Jones/Butcher).
- Dedup festivi D/E e H/I: stesso medico in D+E o H+I nello stesso giorno festivo conta 1 volta, non 2.
- H/I nel riepilogo mostrano solo feriali (`.get("feriali", 0)`), non totali.
- Grafici Plotly: menu a tendina (`st.selectbox`) per scegliere il grafico, non tutti visibili insieme.
- Tab "Per mese" nella tabella riepilogativa storica.
- Default indisponibilità: cambiato a "Usa archivio (privacy)" (`index=2` nel radio widget).
- Auto-carryover da storico: all'importazione del mese, salva `_meta.last_day_night_doctors` nel JSON. Il multiselect carryover nel pannello admin viene pre-compilato con chi ha fatto notte l'ultimo giorno del mese più recente nello storico.
- Nota: la reperibilità (C) è assegnata dal greedy `assign_reperibilita_C`, NON dal CP-SAT. Non può avere soft-constraints storici nel solver.

**Moduli coinvolti:**
| File | Modifica |
|---|---|
| `shift_history.py` | NUOVO — parser dinamico + aggregazione + storage GitHub + easter + valid_doctors |
| `turni_generator.py` | Aggiunto parametro `historical_stats` al solver con soft-constraints (HIST_NIGHT_PENALTY, HIST_FEST_PENALTY, HIST_DEHI_PENALTY) |
| `streamlit_app.py` | Nuova sezione admin "Memoria Storica" + auto-carryover + default archivio |
| `requirements.txt` | Aggiunto `plotly>=5.18.0` |

**TODO futuri (da dove ripartire):**
- [ ] Ri-importare i mesi già caricati nello storico per popolare `_meta.last_day_night_doctors` (i mesi importati prima di questa modifica non hanno `_meta`)
- [ ] Aggiungere soft-constraint storico anche per le domeniche/festivi D/E/H/I (oggi solo notti J e festivi generici)
- [ ] Verificare che il solver usi effettivamente `historical_stats` quando si genera da Streamlit (passaggio del parametro alla pipeline completa)
- [ ] Considerare un riepilogo visivo del carryover nella UI (es. "Da storico: Licordari ha fatto notte il 31/03")
- [ ] Test automatici per `shift_history.py` (parsing, conteggi, edge cases layout diversi)
- [ ] Gestione del caso in cui il mese nello storico non è il mese immediatamente precedente (es. manca un mese intermedio)

## Feature in sviluppo: Gestione Pool Medici da GUI (pool_config)

**Spec completa:** `docs/superpowers/specs/2026-05-07-pool-config-design.md` *(da creare)*

**Decisioni di design confermate (sessione 7 maggio 2026):**

### Approccio: Overlay JSON su YAML (Approccio A)
- `data/pool_config.json` su GitHub sovrascrive pool e quote al momento della generazione
- `Regole_Turni.yml` rimane immutato come template avanzato (spacing, penalità, relief_valves)
- Il merge avviene in `streamlit_app.py` prima di invocare il solver
- La GUI admin (PIN-protetta) legge/scrive solo `pool_config.json`

### Funzionalità previste
1. **Gestione medici** — aggiungere/rimuovere medici, attivo/inattivo
2. **Assegnazione colonne** — per ogni medico: quali colonne può fare normalmente
3. **Festivi** — due toggle per medico: `festivi_diurni` (D/E/H/I) e `festivi_notti` (J nei festivi)
4. **Limiti** — quote min/max globali per colonna + override per singolo medico
5. **Combinazioni same-day** — coppie di servizi che lo stesso medico può coprire nello stesso giorno (es. K+T), estende il meccanismo `df_pair`
6. **Servizi critici** — servizi che non possono mai rimanere scoperti: usa pool primario, fallback su qualsiasi medico attivo se il pool primario è esaurito

### Schema `pool_config.json`
```json
{
  "doctors": {
    "Licordari": {
      "active": true,
      "columns": ["D", "E", "J", "K", "T", "Q"],
      "festivi_diurni": true,
      "festivi_notti": true,
      "column_overrides": { "J": {"quota_min": 2, "quota_max": 4} }
    }
  },
  "column_defaults": {
    "J": {"quota_min": 1, "quota_max": 6},
    "D": {"quota_min": 0, "quota_max": 5}
  },
  "service_combinations": [
    {"columns": ["K", "T"], "same_day": true}
  ],
  "critical_services": ["D", "E", "J", "H"],
  "updated_at": "...",
  "updated_by": "admin"
}
```

### Integrazione solver
- `pool_config.json` viene caricato in `streamlit_app.py` prima della generazione
- La funzione `apply_pool_config(cfg_yaml, pool_config)` in `turni_generator.py` produce il cfg effettivo
- Per i servizi critici: il solver riceve un `emergency_pool` per colonna (tutti i medici attivi) usato solo se il pool primario è esaurito
- Per le combinazioni same-day: il meccanismo `df_pair` viene generalizzato a `service_pairs` con lista configurabile

### File coinvolti
| File | Modifica |
|---|---|
| `streamlit_app.py` | Nuova sezione admin "Gestione Pool" + load/save pool_config.json |
| `turni_generator.py` | Nuova funzione `apply_pool_config()` + generalizzazione `df_pair` → `service_pairs` + logica `critical_services` |
| `pool_config_store.py` | NUOVO — funzioni pure per load/save/validate del pool_config JSON |
| `github_utils.py` | Nessuna modifica (usa `get_file`/`put_file` esistenti) |

### Secrets aggiuntivi (opzionale)
```toml
[github_unavailability]
pool_config_path = "data/pool_config.json"  # default se non presente
```

**Stato:** pianificazione in corso — spec non ancora scritta
