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

**Test** (unittest, include un end-to-end dell'app con `streamlit.testing.v1.AppTest` e GitHub/SMTP finti):
```bash
.venv/bin/python -m unittest discover -s tests -t .
```
Su macOS lanciarla con `caffeinate -i` davanti: se il Mac va in stop durante la suite, AppTest misura il tempo con l'orologio di sistema e al risveglio segnala "AppTest script run timed out after 60(s)" (falso blocco; verifica con `pmset -g log | grep -E " Sleep | Wake "`).

**Sandbox locale** (app vera su una copia dei dati in `.local_sandbox/repo`, nessuna scrittura su GitHub, nessuna mail):
```bash
.venv/bin/streamlit run scripts/run_local_sandbox.py
```

## Architettura

| File | Ruolo |
|---|---|
| `turni_generator.py` | Solver principale: legge template Excel + regole YAML + indisponibilità, costruisce il modello CP-SAT, scrive l'Excel di output |
| `streamlit_app.py` | UI Streamlit: login medico (PIN + OTP email), inserimento indisponibilità, generazione turni per admin |
| `unavailability_store.py` | Funzioni pure per il datastore CSV: parsing, filtraggio, deduplicazione, serializzazione, firme mese e `save_doctor_months()` (salvataggio con controllo conflitti) |
| `unavailability_drafts.py` | Bozze delle modifiche non ancora inviate (autosalvataggio), ripresa al login, riepilogo bozze pendenti per l'admin |
| `unavailability_receipts.py` | Mail di resoconto dopo ogni salvataggio (contenuto dai dati riletti dal server, niente note) |
| `unavailability_calendar.py` + `components/unav_calendar/` | Calendario della pagina medico: logica pura degli eventi (`apply_event`, `calendar_payload`) + componente Streamlit in JS puro |
| `generation_memory.py` | Memoria delle generazioni salvate: uso pregresso per periodi parziali, carryover notti, storico provvisorio |
| `github_utils.py` | Lettura/scrittura via GitHub Contents API (archivia il CSV delle indisponibilità su una repo privata) |
| `xlsx_utils.py` | Genera il file XLSX delle indisponibilità dal CSV usando `unavailability_template.xlsx` |
| `Regole_Turni.yml` | File regole mensile: definizione colonne, pool medici, quote, vincoli, penalità |
| `data/doctor_contacts.yml` | Mappa nome medico → email (usata per gli OTP) |
| `Style_Template.xlsx` | Template di stile opzionale applicato alla generazione di nuovi template Excel mensili |

### Flusso dei dati

1. Le **regole** sono definite in `Regole_Turni.yml` — le lettere di colonna corrispondono ai tipi di turno, ciascuno con pool, quote, vincoli di spaziatura e pesi di penalità.
2. Le **indisponibilità** sono archiviate come CSV per-medico in una repo GitHub privata, uno per file (`data/unavailability/unavail_<slug>.csv`, vedi `_doctor_unavail_path()` in `streamlit_app.py`) — evita conflitti di scrittura concorrente tra medici. `load_store_from_github()` aggrega tutti i file della cartella; se la cartella è vuota usa come fallback legacy il vecchio file unico `data/unavailability_store.csv` (path configurabile via `github_unavailability.path`). Stessa logica a coppie per le **preferenze di disponibilità** (`data/availability/avail_<slug>.csv`, aggregate da `load_avail_store_from_github()`, fallback legacy `data/availability_store.csv`). I medici inseriscono indisponibilità/preferenze via Streamlit; l'admin può anche caricare un file Excel.
3. Il **solver** (`turni_generator.py`) traduce regole + indisponibilità in variabili e vincoli CP-SAT, risolve e compila il workbook openpyxl.
4. **Streamlit** (`streamlit_app.py`) orchestra il tutto per gli utenti web: gestisce auth PIN, OTP via SMTP, lease di sessione per medico (kick-out in caso di login concorrente) e invoca `turni_generator` in-process.

### Salvataggio indisponibilità (regole da non rompere)

- **Salva sostituisce l'intero mese**, quindi ogni editor memorizza la firma del mese da cui è partito (`<rows_key>__base_sig`). `ustore.save_doctor_months()` rifiuta con `MonthConflictError` se nel frattempo il mese è cambiato sul server (altro dispositivo, scheda vecchia, lease scaduto): **mai** riapplicare l'editor su dati freschi dopo un 409. Il lease di sessione da solo NON basta (commit 67c3aca aveva tolto il controllo: una scheda vecchia cancellava in silenzio giorni salvati altrove).
- Mese già uguale all'editor → nessuna scrittura (il doppio tap è un no-op).
- **Bozze**: le modifiche non inviate vengono scritte in `data/unavailability_drafts/draft_<slug>.json` (al cambio, max ogni 15 s, più un fragment `run_every=20`). Al login la bozza viene ripresa (se costruita sui dati attuali) o proposta (se il server è cambiato). Il solver usa SOLO i CSV ufficiali; il pannello "Genera turni" mostra le bozze non inviate.
- Dopo un salvataggio riuscito, baseline/audit/pulizia bozza/mail passano da un'outbox in `session_state` (`_process_unav_outbox`): un doppio tap che interrompe l'esecuzione non fa perdere audit o mail.
- **Email**: un'unica configurazione `[smtp]` nei Secrets per codici PIN (OTP) e resoconti; il pannello admin "📧 Email" ne mostra lo stato e manda una mail di prova. La password resta nei Secrets: il repo dati è pubblico.
- **Mail di resoconto** al medico (email da `doctor_contacts.yml`) + copie da impostazioni (`receipt_cc_emails`, modificabile nel pannello admin, default `utic@polime.it`) e da secrets `[notifications] receipt_cc`. Best-effort: se fallisce il salvataggio resta valido.
- Nessuna riga precompilata negli editor: una riga "1 del mese, Mattina" lasciata lì veniva salvata come indisponibilità vera.
- **Calendario**: niente pulsante "Applica", ogni tocco su una fascia si registra subito in bozza. Il componente rinvia ogni modifica (`edits` con `seq`, più `cid` dell'istanza) finché la pagina non la conferma nel payload (`ack`, in `unav_calendar_ack::<medico>`): Streamlit può unire due tocchi rapidi in un solo rerun e nessuno dei due deve perdersi (`ucal.apply_value`). Anche "Ferie lunghe" passa dalle stesse regole (evento `range` via `_apply_calendar_edit()`), così "Ferie" e "Tutto il giorno" restano esclusive nel giorno. Nel componente l'altezza dell'iframe va misurata sul contenuto (`#root`), mai su `document.documentElement.scrollHeight`: include l'iframe stesso e cresce all'infinito. Nei test AppTest il componente è sostituito da `FakeCalendar` (patch di `declare_component`, stesso protocollo).

### GitHub: limiti, concorrenza, coda (regole da non rompere)

Tutta l'app usa **un solo token**: 5.000 richieste/ora, 80 scritture/minuto e 500/ora (limiti secondari), risposte 403/429 con `retry-after` quando si sforano.

- **Ogni chiamata passa da `github_utils`** (`request()`): retry su rate limit / 5xx / errori di rete rispettando `retry-after` e `x-ratelimit-reset` entro `max_wait`; oltre, `GithubUnavailable(retry_at)` e fail-fast fino a quel momento. Le **scritture sono serializzate nel processo** e distanziate di `WRITE_MIN_INTERVAL` (≤60/min): i commit dell'app non si pestano mai sul branch. 409/422 (sha) e 403 di permessi NON vengono ritentati lì.
- **Lease di sessione in memoria** (`session_leases.py`, `st.cache_resource`): niente più commit a ogni login/heartbeat né lettura ogni 5 s. L'heartbeat (`touch`) non ruba mai il lease a una sessione più recente.
- **Letture in cache**: impostazioni (60 s), pool config per l'elenco medici (120 s), record PIN (5 min), contatti (120 s); ognuna si svuota al salvataggio corrispondente.
- **Bozze in memoria** (`MemoryDraftStore`), copia su GitHub al massimo ogni 120 s come *merge* per mese (un load fallito non può cancellare una bozza su GitHub).
- **Salva** passa da `SaveQueue.save_or_enqueue()` (`unavailability_service.py`): se GitHub resta indisponibile oltre ~25 s il salvataggio va **in coda** (persistita in `/tmp`), un worker thread lo completa e solo allora manda la mail. Salvataggi dell'app e del worker per lo stesso medico sono serializzati (lock per medico). Un mese in coda (o appena registrato dalla coda) da **un'altra sessione** è un conflitto: mai sovrascritto. Se in background risulta un conflitto, non viene applicato e parte una mail "NON salvate" al medico + copie.
- Il worker non chiama mai `st.*`: tutto ciò che gli serve è in `ServiceConfig` catturata nel thread della pagina.
- Audit idempotente: una risposta persa dopo il commit + retry non duplica la riga.
- La coda persiste in `TURNI_QUEUE_PERSIST_PATH` (default `/tmp/turni_unav_save_queue.json`): test e sandbox usano un file proprio, altrimenti riprenderebbero job altrui.
- Test: `tests/fakes.py` simula GitHub a livello HTTP con disturbi (rate limit, 5xx, timeout anche dopo il commit); `tests/test_unavailability_service.py` fa salvataggi concorrenti di 25 medici sotto disturbi; `tests/test_streamlit_unavailability_flow.py` esegue l'app vera con e senza disturbi (AppTest non regge sessioni in thread paralleli: la concorrenza vera è coperta dai test del servizio).

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
path   = "data/unavailability_store.csv"          # fallback legacy, solo se per_doctor_dir è vuota
per_doctor_dir = "data/unavailability"             # default, opzionale se non rinominata
availability_path = "data/availability_store.csv"  # fallback legacy preferenze
per_doctor_avail_dir = "data/availability"         # default, opzionale se non rinominata
drafts_dir = "data/unavailability_drafts"          # default, bozze non inviate

[notifications]
receipt_cc = ["..."]                               # opzionale: copie extra della mail di resoconto

[smtp]
host     = "smtp.gmail.com"
port     = 587
username = "..."
password = "APP_PASSWORD"
from     = "..."
starttls = true
```

### IMPORTANTE: `num_search_workers` del CP-SAT solver

In `solve_with_ortools()` (`turni_generator.py`) tutte le istanze di `cp_model.CpSolver()` — **comprese quelle del retry diagnostico** — usano `num_search_workers = 1` (fino a settembre 2026 cinque solver diagnostici non lo avevano e un mese infeasible bloccava la generazione per sempre; test `test_diagnostic_retry_terminates_when_only_night_constraints_conflict`). **Non aumentare questo valore.** Con `num_search_workers > 1` (testato con 8) il `max_time_in_seconds` non viene rispettato: il `solver.Solve()` può bloccarsi indefinitamente (osservato >125s contro un limite di 30s), sia in locale (macOS arm64) sia su Streamlit Cloud — bug della ricerca multi-thread di OR-Tools su questa combinazione di piattaforma/versione (`ortools` 9.15.6755), non un problema delle regole/vincoli. Con `num_search_workers = 1` il timeout è sempre rispettato (verificato: 30.0s esatti). Se si aggiorna `ortools` in futuro e si vuole ritentare il multi-thread, verificare prima con un test di carico che il timeout venga rispettato in modo affidabile.

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

**Spec completa:** `docs/superpowers/specs/2026-05-07-pool-config-design.md` ✅ rev. 2026-05-07b  
**Mockup interattivo:** `docs/superpowers/mockup_pool_config.html` ✅ (aprire in browser)  
**Stato:** ✅ IMPLEMENTATA E DEPLOYATA (commit df8933e → 6a2962a)

### Approccio: Overlay JSON su YAML (Approccio A)
- `data/pool_config.json` su GitHub sovrascrive pool e quote al momento della generazione
- `Regole_Turni.yml` rimane immutato come template avanzato (spacing, penalità, relief_valves)
- Il merge avviene tramite `apply_pool_config(cfg_yaml, pool_config)` in `turni_generator.py`
- La GUI admin (PIN-protetta) legge/scrive solo `pool_config.json`

### Funzionalità — design finalizzato
1. **Gestione medici** — aggiungere/rimuovere medici, attivo/inattivo, festivi diurni/notturni, reperibilità, universitario+ratio
2. **Assegnazione colonne** — per ogni medico: griglia di tutti i servizi con toggle
3. **Quote** — `monthly_target` per colonna (default per tutti) + override per singolo medico con tipo `fixed` (sempre esattamente N) o `max` (mai più di N). Nessun min/max separato.
4. **Notti weekend** — override `weekend_nights: false` per J (es. Calabrò escluso sab/dom)
5. **Combinazioni same-day** — servizi che lo stesso medico può coprire nello stesso giorno con modalità `always`/`fallback`/`preferred`
6. **Servizi critici** — mai scoperti: fallback `any` (qualsiasi medico attivo) o lista esplicita

### Logica quote notti J — CONFERMATA
| Medico | Override | Comportamento |
|---|---|---|
| Zito, Dattilo | `quota: 2, type: max` | Mai più di 2 — vincolo hard; ratio universitario già applicato |
| De Gregorio | J non in `columns` | Escluso dalle notti (non nel pool) |
| Calabrò | `weekend_nights: false` | Notti feriali sì, sab/dom no |
| Licordari, Colarusso e tutti gli altri | nessuno | Target flessibile 2; fanno 3 a rotazione se necessario (aggiornato 23 giugno 2026: rimossa quota fissa 3 per Licordari/Colarusso) |

### Regole chiave chiarite
- **J vale 2 turni per tutti** (ospedalieri e universitari) → `counts_as: 2` in `column_settings`
- **`university_doctor.ratio`** → il solver lo applica già al workload mensile Mon-Sat esclusi festivi; la GUI espone solo il numero
- **K+T**: mode `always` (stesso medico obbligatorio ogni giorno — già fanno le stesse cose fisicamente)
- **Q+R**: mode `fallback` (separati se possibile, accoppiati solo se pool esaurito)

### Schema `pool_config.json` (v1 — aggiornato 7 maggio 2026)
```json
{
  "schema_version": 1,
  "doctors": {
    "Licordari": {
      "active": true,
      "columns": ["D", "E", "J", "K", "T", "Q"],
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
      "columns": ["D", "E", "J"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": { "ratio": 0.6 },
      "column_overrides": {}
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
    "Grimaldi": {
      "active": true,
      "columns": ["D", "F"],
      "festivi_diurni": false,
      "festivi_notti": false,
      "excluded_from_reperibilita": true,
      "university_doctor": null,
      "column_overrides": {}
    }
  },
  "column_settings": {
    "J": { "monthly_target": 2, "spacing_min_days": 5, "balance_weight": 300, "counts_as": 2 },
    "D": { "monthly_target": null, "spacing_min_days": 0, "balance_weight": 200, "counts_as": 1 },
    "C": { "monthly_target": null, "spacing_min_days": 0, "balance_weight": 200, "counts_as": 1 }
  },
  "_note_column_settings": "monthly_target = quota di default per il turno J per tutti i medici senza override. counts_as=2 per J vale per tutti.",
  "_note_column_overrides": "Tre tipi di override quota per colonna J: fixed (sempre esattamente N), max (mai più di N, ma può fare meno), nessun override (default flessibile, può fare +1 in rotazione se necessario). weekend_nights:false esclude sab/dom.",
  "_note_notti_logic": "Zito e Dattilo: max=2 (MAI 3). Tutti gli altri (incl. Licordari e Colarusso): default flessibile (di solito 2, fanno 3 a turno se serve).",
  "service_combinations": [
    {
      "columns": ["K", "T"],
      "same_day": true,
      "mode": "always"
    },
    {
      "columns": ["Q", "R"],
      "same_day": true,
      "mode": "fallback"
    }
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

### Integrazione solver
- `pool_config.json` viene caricato in `streamlit_app.py` prima della generazione
- La funzione `apply_pool_config(cfg_yaml, pool_config)` in `turni_generator.py` produce il cfg effettivo
- Per i servizi critici: il solver riceve un `emergency_pool` per colonna (tutti i medici attivi) usato solo se il pool primario è esaurito
- Per le combinazioni same-day: il meccanismo `df_pair` viene generalizzato a `service_pairs` con tre modalità:
  - `always`: vincolo HARD — stesso medico obbligatorio per entrambe le colonne nello stesso giorno
  - `fallback`: vincolo SOFT ad alta penalità — il solver preferisce medici separati, li accoppia solo se pool esaurito (comportamento attuale di `enable_kt_share` in `relief_valves`)
  - `preferred`: vincolo SOFT a bassa penalità — il solver preferisce accoppiarli ma non è obbligatorio
- Per `weekend_nights: false` in `column_overrides.J`: il medico viene escluso dal pool J nei giorni sabato e domenica (sostituisce `weekend_excluded_doctors` hardcoded nel YAML)

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
