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
| `festivi_store.py` | Turni festivi pre-assegnati (sorteggio + assegnazione manuale) della pagina **Admin → Festivi**: festività/prefestivi del periodo, estrazione bilanciata, controlli, conversione per il solver. File `data/turni_festivi.yml` nel repo **dati** (non si modifica più a mano); quello nel repo codice resta solo come copia iniziale e per la CLI |
| `doctor_contacts.py` | Email dei medici in `doctor_contacts.yml` (codici PIN + mail di resoconto), modificate dal pannello pool → Medici |
| `unavailability_calendar.py` + `components/unav_calendar/` | Calendario della pagina medico: logica pura degli eventi (`apply_event`, `calendar_payload`) + componente Streamlit in JS puro |
| `doctor_registry.py` + `doctor_admin_ui.py` | Schede medico (unica fonte di chi fa cosa) + pagina Admin → Medici |
| `generation_memory.py` | (non più usato dall'app) vecchia memoria delle generazioni salvate |
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
- **Preferenze ("Preferirei") = stesse regole**: viaggiano nello stesso `SaveJob` (`pref_entries_by_month`, `pref_base_signatures`, file `data/availability/avail_<slug>.csv`) con firma che include la priorità (`with_priority=True`), conflitti, coda, bozza, audit (`save_pref`) e mail. I conflitti di entrambi i file sono controllati prima della prima scrittura; `job.landed` ricorda cosa è già stato scritto se GitHub cade tra i due file, così la mail finale li elenca entrambi. Fino a ottobre 2026 passavano da un percorso separato (`save_doctor_availability_with_retry`, rimosso): niente bozza, niente coda, una scheda vecchia sovrascriveva le preferenze salvate da un altro dispositivo e senza modifiche alle indisponibilità non compariva alcun messaggio.
- **Bozze**: le modifiche non inviate vengono scritte in `data/unavailability_drafts/draft_<slug>.json` (al cambio, max ogni 15 s, più un fragment `run_every=20`). Al login la bozza viene ripresa (se costruita sui dati attuali) o proposta (se il server è cambiato). Il solver usa SOLO i CSV ufficiali; il pannello "Genera turni" mostra le bozze non inviate.
- Dopo un salvataggio riuscito, baseline/audit/pulizia bozza/mail passano da un'outbox in `session_state` (`_process_unav_outbox`): un doppio tap che interrompe l'esecuzione non fa perdere audit o mail.
- **Email**: un'unica configurazione per codici PIN (OTP) e resoconti = Secrets `[smtp]` sovrascritti dal pannello admin "📧 Email" (`email_settings.py`, file `data/email_settings.json`). Il repo dati è pubblico: la password (password per le app di Google) sta lì **solo cifrata** (Fernet, chiave ricavata dal token GitHub dei Secrets), viene provata col server prima di salvarla e non viene mai rimostrata. Se cambia il token GitHub la password salvata non è più leggibile: l'app torna ai Secrets e il pannello chiede di reinserirla. Nella sandbox locale l'SMTP è finto (stampa in console).
- **Mail di resoconto** al medico (email da `doctor_contacts.yml`, modificabile nella scheda del medico in Admin → Medici, che è l'unico punto di modifica) + copie da impostazioni (`receipt_cc_emails`, modificabile nel pannello admin, default `utic@polime.it`) e da secrets `[notifications] receipt_cc`. Best-effort: se fallisce il salvataggio resta valido.
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
email_settings_path = "data/email_settings.json"   # default, credenziali mail dal pannello admin (password cifrata)

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

### Solver a fasi: mai perdere una soluzione già trovata

`solve_with_ortools()` risolve in fasi successive (copertura → condivisioni K+T → servizi pesanti → **equità delle notti** → obiettivo finale), ognuna con un vincolo "non peggio di prima". Ogni fase parte dalla soluzione della precedente (`_hint_current_solution`) e, se il tempo scade senza nuova soluzione, la riusa (`_solve_keeping_previous`). Fino a novembre 2026 l'ultima fase ripartiva da zero: dicembre 2026 (molte assenze a Natale) trovava l'ottimo alla fase 3, poi l'ultima fase restava UNKNOWN dopo 60 s e l'intero mese usciva "INFEASIBLE" e vuoto. Un'assegnazione fissa/sorteggio (D, H, J) di un medico non attivo è "non applicabile": il turno resta **vuoto e giallo** (mai dato al pool né al pool di emergenza) finché l'admin non lo riassegna nella pagina Festivi. Il sorteggio prevale sulle altre regole (le indisponibilità in conflitto vengono ignorate, `_strip_festivi_unavailability`); una notte assegnata in un giovedì di J vuota viene creata.

**Equità delle notti** (fase `night_fair_terms`: fascia min/max dei medici senza quota, differenza max-min, quote J "fisse" della scheda, tutte a 200M): ha una fase sua perché su Streamlit Cloud (CPU lenta) l'ultima fase si ferma presto e l'equità, con peso basso e mischiata a tutto il resto, restava indietro (dicembre 2026: Allegra 1 notte, Colarusso 4; riprodotto in locale con `solver_max_time_seconds` 4). Le protezioni dei servizi piccoli (`next_day_tiny_pool_night_penalty`, ora 10M) non devono mai pesare più dell'equità.

### IMPORTANTE: `num_search_workers` del CP-SAT solver

In `solve_with_ortools()` (`turni_generator.py`) tutte le istanze di `cp_model.CpSolver()` — **comprese quelle del retry diagnostico** — usano `num_search_workers = 1` (fino a settembre 2026 cinque solver diagnostici non lo avevano e un mese infeasible bloccava la generazione per sempre; test `test_diagnostic_retry_terminates_when_only_night_constraints_conflict`). **Non aumentare questo valore.** Con `num_search_workers > 1` (testato con 8) il `max_time_in_seconds` non viene rispettato: il `solver.Solve()` può bloccarsi indefinitamente (osservato >125s contro un limite di 30s), sia in locale (macOS arm64) sia su Streamlit Cloud — bug della ricerca multi-thread di OR-Tools su questa combinazione di piattaforma/versione (`ortools` 9.15.6755), non un problema delle regole/vincoli. Con `num_search_workers = 1` il timeout è sempre rispettato (verificato: 30.0s esatti). Se si aggiorna `ortools` in futuro e si vuole ritentare il multi-thread, verificare prima con un test di carico che il timeout venga rispettato in modo affidabile.

### Keep-alive (GitHub Actions)

`.github/workflows/keep_awake_selenium.yml` fa un ping all'app Streamlit ogni 3 ore tramite Selenium per evitare l'hibernation. Richiede il secret Actions `STREAMLIT_APP_URL`.

## Convenzioni importanti

- I nomi dei medici nelle schede (Admin → Medici) devono corrispondere **esattamente** ai nomi nei file di indisponibilità e in `doctor_contacts.yml`.
- Valori ammessi per `Fascia`: `Mattina`, `Pomeriggio`, `Notte`, `Diurno` (= Mattina + Pomeriggio), `Tutto il giorno`.
- Le lettere di colonna (C, D, E, …) nel YAML corrispondono direttamente alle colonne Excel del template.
- I pool e i nomi ancora presenti in `Regole_Turni.yml` servono SOLO alla migrazione iniziale e alla CLI: con le schede sul repo dati vengono ignorati. Per reinserire un medico (es. D'Angelo) o escluderlo si usa la sua scheda (Attivo + servizi).
- La sezione `relief_valves` definisce fallback ad alta penalità per evitare l'infeasibility (es. permettere una colonna vuota a costo elevato anziché fallire).

## Schede medico (`doctor_registry.py`) — unica fonte di chi fa cosa

**Regola da non rompere:** il solver prende i medici, i servizi che possono fare e ogni regola personale SOLO dalle schede in `data/doctor_registry.json` (repo dati). `doctor_registry.to_solver_cfg(cfg_yaml, registry)` parte dal YAML **svuotato di ogni nome** (`strip_names`: pool, quote, esclusioni, `friday_required_doctor`, `tuesday_fixed`, `pattern_doc1/2`, `university_doctors`, `absolute_exclusions`…) e lo riempie dalle schede; `cfg["doctors"]` è l'elenco attivi usato da `collect_doctors`. Test `test_no_doctor_can_come_from_outside_the_cards`. Nel solver non ci sono più nomi scritti nel codice (Grimaldi/Calabrò fuori da H, Crea venerdì V, Cimino U, De Gregorio I, Allegra V+U, never_in_J di default…): tutto arriva da cfg.

- **Scheda**: attivo, reperibilità, festivi diurni/notti, quota festivi diurni (`festivi_quota`, es. Manganaro 2 = causa delle "troppe domeniche"), no sabato diurno, esente 2 weekend liberi, universitario (ratio, notte doppia), mai notti (nemmeno sorteggio), `lunedi_V_U_con`, note, e `services` per servizio (`D`, `F`, `EG` = E+G sempre insieme, `H`…`AC`): livello `pool` (assegnabile) o `riserva` (solo se nessuno del pool è disponibile quel giorno: sostituisce i vecchi fallback "qualsiasi medico"), `quota` {fixed|max|min, value, target}, opzioni J (notti weekend, min/max weekend), V (`doppio` titolare/secondo), W (`giorni_fissi`), AB (`sabato`). Coppia D/F = i 2 medici con D a livello pool (alternanza 3+3 solo se sono esattamente 2).
- **Quote generiche** (`pool_quota_overrides`) su D/E/G/H/I contano solo i giorni feriali (i festivi hanno `festivi_quota`). Quota J "fixed" = quasi obbligatoria e mai superata (ex Zito/Dattilo/Calabrò). Recupero: le sue quote T/Q/W/L tornano nelle chiavi `recupero_*` del YAML.
- **Migrazione** `migrate(pool_config_v1, yaml)`: costruisce le schede dalla configurazione precedente senza perdere niente (comprese le riserve che il solver usava di nascosto); `active` della scheda vince sulla vecchia `absolute_exclusions` del YAML. Finché il file non esiste l'app mostra le schede migrate e il primo salvataggio le scrive tutte. Verificato: per apr/giu/ago/nov/dic 2026 e gen/mar 2027, con indisponibilità e sorteggi reali, i candidati di ogni casella sono identici a prima (unica differenza: D'Angelo ora attivo come impostato dall'admin).
- **Pagina Admin → Medici** (`doctor_admin_ui.py`): un medico alla volta in un form; chiavi dei widget con nome medico + versione (prima, `data_editor` a righe dinamiche + chiavi non legate al medico avevano salvato su Allegra le colonne scelte per D'Angelo, 18/11/2026); `update_registry` rilegge il file, rifiuta se quella scheda è cambiata nel frattempo (`CardChanged`) e sostituisce solo quella scheda. Anche le impostazioni per servizio (target, distanze, "vale come", K+T, indispensabili) sono lì. `pool_config.json` non si usa più se non come sorgente della migrazione.
- CLI/GUI tkinter: `load_solver_cfg(rules)` = schede migrate dal YAML.

## Bilanciamento con i turni definitivi (`shift_history.py`, pagina Admin → Storico turni)

- Si caricano gli Excel **definitivi** (anche più file, anche a cavallo di due mesi). `merge_upload` salva per giorno (`_days` nel mese): un nuovo file sostituisce solo i giorni che contiene, le statistiche si ricalcolano dai giorni (niente doppi conteggi se arriva prima 1/12-7/1 e poi gennaio). `aggregate_multi_month` ignora le chiavi che iniziano con `_`.
- Le **generazioni salvate non contano più** (rimossa dall'app la "memoria generazioni"; `generation_memory.py` resta solo per i suoi test). Notte del giorno prima del periodo: `night_doctors_on(history, data)`.
- Solver: `history_fair_targets` per notti, notti sab/dom, domeniche/festivi di giorno su **tutti i mesi caricati**: storico → tasso per mese di presenza (chi non ha storico = media del gruppo), poi livello comune "a riempimento"; penalità solo sopra l'obiettivo (`HISTORY_BALANCE_WEIGHTS`, sotto l'equità del periodo e le quote, sopra le preferenze minori). Chi ha quota fissa è escluso e la sua quota è tolta dal totale.
- Reperibilità (`assign_reperibilita_C`): mai il giorno della propria notte né il giorno dopo (non si rilassa mai: meglio C vuota e gialla); la reperibilità in più va a chi ne ha fatte meno nello storico (per mese, media per chi è nuovo).
- Output: un turno obbligatorio rimasto scoperto è giallo anche se indispensabile.
- `github_utils.put_file`: su 409/412/422 rilegge il file; se contiene già esattamente il testo scritto (commit riuscito, risposta persa) è un successo.
