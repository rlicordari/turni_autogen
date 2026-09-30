# -*- coding: utf-8 -*-
"""Unavailability saves that must not get lost.

Everything here is independent from Streamlit (no st.* calls), so the same code
runs in the page and in the background worker:

- save_now(): conflict-safe save of whole doctor-months (see ustore.save_doctor_months)
  through the GitHub gateway (retries, pacing, rate limits);
- write_audit(): idempotent append to the monthly audit CSV (a lost response
  followed by a retry never duplicates a row);
- send_receipt() / notify_not_saved(): emails to the doctor + copies;
- SaveQueue: when GitHub stays unavailable, the save is queued (and persisted on
  local disk) and a worker thread completes it later, then sends the receipt.
  A queued save that turns out to conflict is NEVER applied: the doctor is told.
"""

from __future__ import annotations

import copy
import csv
import datetime as dt
import io
import json
import re
import smtplib
import threading
import time
import uuid
import weakref
from dataclasses import dataclass, field
from email.message import EmailMessage
from pathlib import Path
from typing import Callable, Optional

import requests

import github_utils
import unavailability_receipts as receipts
import unavailability_store as ustore

AUDIT_FIELDS = [
    "ts_utc",
    "doctor",
    "month",
    "action",
    "before_count",
    "after_count",
    "added_count",
    "removed_count",
    "note_changed_count",
    "details_json",
    "app_build",
]

QUEUE_MAX_WAIT = 60.0       # a background attempt may wait this long for GitHub

_ALL_QUEUES = weakref.WeakSet()


def stop_all_workers() -> None:
    """Stop every queue worker thread of this process (tests / shutdown)."""
    for q in list(_ALL_QUEUES):
        q.stop_worker()
QUEUE_MAX_ATTEMPTS = 30     # unexpected errors: give up (and tell the doctor) after this


@dataclass
class ServiceConfig:
    github: dict
    smtp: dict = field(default_factory=dict)
    receipt_cc: list = field(default_factory=list)
    app_build: str = ""


def doctor_slug(doctor: str) -> str:
    s = (doctor or "").strip().lower()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-z0-9_\-]", "", s)
    return s or "doctor"


def utc_now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _gh(cfg: ServiceConfig) -> dict:
    g = dict(cfg.github or {})
    g.setdefault("branch", "main")
    return g


def unavail_path(cfg: ServiceConfig, doctor: str) -> str:
    base = str(_gh(cfg).get("per_doctor_dir") or "data/unavailability").rstrip("/")
    return f"{base}/unavail_{doctor_slug(doctor)}.csv"


def audit_path(cfg: ServiceConfig, month_key: str) -> str:
    base = str(_gh(cfg).get("audit_dir") or "data/unavailability_audit").rstrip("/")
    return f"{base}/unavailability_audit_{month_key}.csv"


def is_conflict_error(err: Exception) -> bool:
    """HTTP error meaning "the file changed since you read it" (sha mismatch)."""
    if isinstance(err, ustore.ShaConflictError):
        return True
    if not isinstance(err, requests.HTTPError):
        return False
    resp = getattr(err, "response", None)
    code = getattr(resp, "status_code", None)
    if code in (409, 412):
        return True
    if code == 422:
        try:
            msg = str((resp.json() or {}).get("message", "")).lower()
        except Exception:
            msg = str(getattr(resp, "text", "") or "").lower()
        return "sha" in msg
    return False


# ── GitHub I/O ───────────────────────────────────────────────────────────────

def load_doctor_rows(cfg: ServiceConfig, doctor: str, *, max_wait: Optional[float] = None):
    g = _gh(cfg)
    gf = github_utils.get_file(g["owner"], g["repo"], unavail_path(cfg, doctor), g["token"], g["branch"], max_wait=max_wait)
    if gf is None:
        return [], None
    return ustore.load_store(gf.text), gf.sha


def _put_rows_fn(cfg: ServiceConfig, doctor: str, message: str, *, max_wait: Optional[float] = None):
    def _save(rows, sha):
        g = _gh(cfg)
        resp = github_utils.put_file(
            g["owner"], g["repo"], unavail_path(cfg, doctor), g["token"], message, ustore.to_csv(rows),
            branch=g["branch"], sha=sha, max_wait=max_wait,
        )
        resp = resp if isinstance(resp, dict) else {}
        return {
            "content_sha": (resp.get("content") or {}).get("sha"),
            "commit_sha": (resp.get("commit") or {}).get("sha"),
        }
    return _save


# ── Drafts on GitHub (copy of the in-memory drafts) ─────────────────────────

def draft_path(cfg: ServiceConfig, doctor: str) -> str:
    import unavailability_drafts as udrafts
    base = str(_gh(cfg).get("drafts_dir") or udrafts.DRAFTS_DIR_DEFAULT).rstrip("/")
    return f"{base}/draft_{doctor_slug(doctor)}.json"


def _load_draft_with_sha(cfg: ServiceConfig, doctor: str, *, max_wait: Optional[float] = None):
    import unavailability_drafts as udrafts
    g = _gh(cfg)
    gf = github_utils.get_file(g["owner"], g["repo"], draft_path(cfg, doctor), g["token"], g["branch"], max_wait=max_wait)
    if gf is None:
        return udrafts.empty_draft(doctor), None
    draft = udrafts.from_text(gf.text)
    draft["doctor"] = doctor
    return draft, gf.sha


def load_draft(cfg: ServiceConfig, doctor: str, *, max_wait: Optional[float] = None) -> dict:
    return _load_draft_with_sha(cfg, doctor, max_wait=max_wait)[0]


def update_draft(cfg: ServiceConfig, doctor: str, apply_fn, *, max_wait: Optional[float] = None, max_retries: int = 8) -> dict:
    """Read-modify-write of the doctor's draft file, retried on SHA conflicts."""
    import unavailability_drafts as udrafts
    g = _gh(cfg)
    for attempt in range(max_retries):
        draft, sha = _load_draft_with_sha(cfg, doctor, max_wait=max_wait)
        new_draft = apply_fn(draft)
        new_draft["doctor"] = doctor
        if sha is None and not new_draft.get("months"):
            return new_draft  # nothing to store
        try:
            github_utils.put_file(
                g["owner"], g["repo"], draft_path(cfg, doctor), g["token"],
                f"Draft unavailability: {doctor} ({utc_now_iso()})", udrafts.to_text(new_draft),
                branch=g["branch"], sha=sha, max_wait=max_wait,
            )
            return new_draft
        except Exception as e:
            if is_conflict_error(e) and attempt < max_retries - 1:
                github_utils._sleep(min(3.0, 0.2 * (2 ** attempt)))
                continue
            raise
    return {}


# ── Save job ─────────────────────────────────────────────────────────────────

@dataclass
class SaveJob:
    job_id: str
    doctor: str
    doctor_email: str
    entries_by_month: dict           # {(yy, mm): [(date, shift, note), ...]}
    base_signatures: dict            # {(yy, mm): [(date_iso, shift, note), ...]}
    by_admin: bool = False
    created_at: float = 0.0
    attempts: int = 0
    next_attempt_at: float = 0.0
    last_error: str = ""
    origins: dict = field(default_factory=dict)   # {(yy, mm): editor/session id}


def make_job(
    *,
    doctor: str,
    doctor_email: str,
    entries_by_month: dict,
    base_signatures: dict,
    by_admin: bool = False,
    created_at: Optional[float] = None,
    origin: str = "",
) -> SaveJob:
    """`origin` identifies the editor session: a month waiting in the queue (or
    just landed from it) can be continued only by the same session."""
    return SaveJob(
        job_id=str(uuid.uuid4()),
        doctor=doctor,
        doctor_email=str(doctor_email or ""),
        entries_by_month={tuple(k): list(v) for k, v in entries_by_month.items()},
        base_signatures={tuple(k): ustore.normalize_signature(v) for k, v in base_signatures.items()},
        by_admin=by_admin,
        created_at=time.time() if created_at is None else created_at,
        origins={tuple(k): str(origin) for k in entries_by_month},
    )


@dataclass
class SaveResult:
    job: SaveJob
    outcome: ustore.SaveOutcome
    saved_at: str


def save_now(cfg: ServiceConfig, job: SaveJob, *, max_wait: Optional[float] = None) -> SaveResult:
    """Save the job's months now. Raises GithubUnavailable or MonthConflictError."""
    saved_at = utc_now_iso()
    prefix = "Admin update" if job.by_admin else "Update"
    outcome = ustore.save_doctor_months(
        load_fn=lambda: load_doctor_rows(cfg, job.doctor, max_wait=max_wait),
        save_fn=_put_rows_fn(cfg, job.doctor, f"{prefix} unavailability: {job.doctor} ({saved_at})", max_wait=max_wait),
        doctor=job.doctor,
        entries_by_month=job.entries_by_month,
        base_signatures=job.base_signatures,
        updated_at=saved_at,
        is_conflict=is_conflict_error,
        sleep_fn=lambda s: github_utils._sleep(s),
    )
    return SaveResult(job=job, outcome=outcome, saved_at=saved_at)


# ── Audit (idempotent) ───────────────────────────────────────────────────────

def _audit_line(row: dict) -> str:
    buf = io.StringIO()
    csv.DictWriter(buf, fieldnames=AUDIT_FIELDS).writerow({k: row.get(k, "") for k in AUDIT_FIELDS})
    return buf.getvalue().strip("\r\n")


def append_audit_row(cfg: ServiceConfig, month_key: str, row: dict, *, max_wait: Optional[float] = None, max_retries: int = 10) -> None:
    g = _gh(cfg)
    path = audit_path(cfg, month_key)
    line = _audit_line(row)
    header = ",".join(AUDIT_FIELDS)
    for attempt in range(max_retries):
        gf = github_utils.get_file(g["owner"], g["repo"], path, g["token"], g["branch"], max_wait=max_wait)
        text = gf.text if gf else ""
        if line in text:
            return  # already written (e.g. response of a previous attempt was lost)
        lines = [l for l in text.splitlines() if l.strip()]
        if not lines or lines[0].strip() != header:
            lines = [header] + (lines[1:] if lines else [])
        new_text = "\n".join(lines + [line]) + "\n"
        try:
            github_utils.put_file(
                g["owner"], g["repo"], path, g["token"],
                f"Audit unavailability {month_key}: {row.get('doctor', '')}", new_text,
                branch=g["branch"], sha=gf.sha if gf else None, max_wait=max_wait,
            )
            return
        except Exception as e:
            if is_conflict_error(e) and attempt < max_retries - 1:
                github_utils._sleep(min(5.0, 0.2 * (2 ** attempt)))
                continue
            raise


def write_audit(cfg: ServiceConfig, result: SaveResult, *, action: Optional[str] = None, max_wait: Optional[float] = None) -> list:
    """One audit row per changed month. Returns warnings (never raises)."""
    warnings = []
    action = action or ("admin_save" if result.job.by_admin else "save")
    for mk, diff in sorted((result.outcome.diffs or {}).items()):
        row = {
            "ts_utc": result.saved_at,
            "doctor": result.job.doctor,
            "month": mk,
            "action": action,
            "before_count": diff.get("before_count", 0),
            "after_count": diff.get("after_count", 0),
            "added_count": diff.get("added_count", 0),
            "removed_count": diff.get("removed_count", 0),
            "note_changed_count": diff.get("note_changed_count", 0),
            "details_json": json.dumps(diff.get("details", {}), ensure_ascii=False),
            "app_build": cfg.app_build,
        }
        try:
            append_audit_row(cfg, mk, row, max_wait=max_wait)
        except Exception as e:
            warnings.append(f"Audit log non aggiornato per {mk}: {e}")
    return warnings


# ── Email ────────────────────────────────────────────────────────────────────

def email_configured(smtp_cfg: dict) -> bool:
    return bool((smtp_cfg or {}).get("host") and (smtp_cfg or {}).get("from"))


def send_email(smtp_cfg: dict, to: list, cc: list, subject: str, body: str) -> None:
    if not email_configured(smtp_cfg):
        raise RuntimeError("Invio email non configurato (smtp).")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_cfg.get("from")
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.set_content(body)
    with smtplib.SMTP(str(smtp_cfg.get("host") or ""), int(smtp_cfg.get("port") or 587), timeout=20) as s:
        s.ehlo()
        if bool(smtp_cfg.get("starttls", True)):
            s.starttls()
            s.ehlo()
        if smtp_cfg.get("username") and smtp_cfg.get("password"):
            s.login(str(smtp_cfg["username"]), str(smtp_cfg["password"]))
        s.send_message(msg)


def _mask(addr: str) -> str:
    name, _, dom = str(addr).partition("@")
    return f"{name[:2]}***@{dom}" if dom else "***"


def send_receipt(cfg: ServiceConfig, result: SaveResult) -> str:
    """Email what the server now holds for the changed months. Never raises."""
    try:
        if not result.outcome.changed:
            return ""
        if not email_configured(cfg.smtp):
            return "⚠️ Mail di resoconto non inviata: SMTP non configurato."
        to, cc = receipts.recipients(result.job.doctor_email, cfg.receipt_cc)
        if not to:
            return "⚠️ Mail di resoconto non inviata: nessun destinatario configurato."
        doctor = result.job.doctor
        rows_by_month = {
            mk: ustore.filter_doctor_month(result.outcome.verified_rows, doctor, int(mk[:4]), int(mk[5:7]))
            for mk in result.outcome.diffs
        }
        subject, body = receipts.build_receipt(
            doctor, rows_by_month, result.outcome.diffs,
            saved_at_utc=result.saved_at, commit_sha=result.outcome.commit_sha, by_admin=result.job.by_admin,
        )
        send_email(cfg.smtp, to, cc, subject, body)
        missing = "" if result.job.doctor_email else " (email del medico mancante in anagrafica)"
        return "📧 Resoconto inviato a " + ", ".join(_mask(a) for a in to + cc) + missing
    except Exception as e:
        return f"⚠️ Mail di resoconto non inviata: {type(e).__name__}: {e}"


def notify_not_saved(cfg: ServiceConfig, job: SaveJob, reason: str) -> str:
    """Tell the doctor (and copies) that a queued save was NOT applied. Never raises."""
    try:
        if not email_configured(cfg.smtp):
            return "SMTP non configurato"
        to, cc = receipts.recipients(job.doctor_email, cfg.receipt_cc)
        if not to:
            return "nessun destinatario"
        lines = [
            f"ATTENZIONE: le indisponibilità di {job.doctor} inviate il "
            f"{dt.datetime.fromtimestamp(job.created_at, receipts.LOCAL_TZ):%d/%m/%Y %H:%M} NON sono state registrate.",
            "",
            f"Motivo: {reason}",
            "",
            "Rientra nell'app, controlla le indisponibilità registrate e reinserisci quelle mancanti.",
            "",
            "Modifiche NON applicate:",
        ]
        for (yy, mm), entries in sorted(job.entries_by_month.items()):
            lines.append(f"  {receipts.month_label(f'{yy:04d}-{mm:02d}')}:")
            for d, sh, _note in sorted(entries):
                lines.append(f"    {d:%d/%m}  {sh}")
            if not entries:
                lines.append("    (nessuna: il mese doveva restare vuoto)")
        send_email(cfg.smtp, to, cc, f"⚠️ Indisponibilità NON salvate – {job.doctor}", "\n".join(lines))
        return "notificato"
    except Exception as e:
        return f"notifica fallita: {e}"


# ── Queue ────────────────────────────────────────────────────────────────────

def _job_to_json(job: SaveJob) -> dict:
    def months(d):
        return {
            f"{int(yy):04d}-{int(mm):02d}": [
                [x[0].isoformat() if isinstance(x[0], dt.date) else str(x[0]), str(x[1]), str(x[2])]
                for x in items
            ]
            for (yy, mm), items in d.items()
        }
    return {
        "job_id": job.job_id,
        "doctor": job.doctor,
        "doctor_email": job.doctor_email,
        "entries_by_month": months(job.entries_by_month),
        "base_signatures": months(job.base_signatures),
        "by_admin": job.by_admin,
        "created_at": job.created_at,
        "attempts": job.attempts,
        "next_attempt_at": job.next_attempt_at,
        "last_error": job.last_error,
        "origins": {f"{int(yy):04d}-{int(mm):02d}": o for (yy, mm), o in job.origins.items()},
    }


def _job_from_json(data: dict) -> SaveJob:
    def key(mk):
        return int(mk[:4]), int(mk[5:7])
    return SaveJob(
        job_id=data["job_id"],
        doctor=data["doctor"],
        doctor_email=data.get("doctor_email", ""),
        entries_by_month={
            key(mk): [(dt.date.fromisoformat(x[0]), x[1], x[2]) for x in items]
            for mk, items in (data.get("entries_by_month") or {}).items()
        },
        base_signatures={
            key(mk): [tuple(x) for x in items] for mk, items in (data.get("base_signatures") or {}).items()
        },
        by_admin=bool(data.get("by_admin")),
        created_at=float(data.get("created_at") or 0),
        attempts=int(data.get("attempts") or 0),
        next_attempt_at=float(data.get("next_attempt_at") or 0),
        last_error=str(data.get("last_error") or ""),
        origins={key(mk): str(o) for mk, o in (data.get("origins") or {}).items()},
    )


class SaveQueue:
    """Saves waiting for GitHub. One pending job per doctor (newer months win).

    Every save of a doctor (interactive or background) runs under that doctor's
    lock, so the worker can never apply an older queued version after, or in
    parallel with, a newer save made from the page.
    """

    def __init__(
        self,
        cfg: ServiceConfig,
        *,
        persist_path: Optional[Path] = None,
        clock: Optional[Callable[[], float]] = None,
        on_saved: Optional[Callable[[SaveResult], None]] = None,
    ):
        self.cfg = cfg
        # Same time source as the gateway: retry_at values come from it.
        self._clock = clock or (lambda: github_utils._now())
        self._persist_path = Path(persist_path) if persist_path else None
        self._on_saved = on_saved
        self._lock = threading.RLock()
        self._doctor_locks: dict = {}
        self._jobs: dict = {}
        self._results: dict = {}
        self._landed: dict = {}   # doctor -> {(yy, mm): (origin, base_before, signature_after)}
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._load()
        _ALL_QUEUES.add(self)

    # persistence (best-effort: survives a process restart in the same container)
    def _load(self) -> None:
        if not self._persist_path or not self._persist_path.exists():
            return
        try:
            data = json.loads(self._persist_path.read_text(encoding="utf-8"))
            for item in data.get("jobs", []):
                job = _job_from_json(item)
                self._jobs[job.doctor] = job
        except Exception:
            pass

    def _save(self) -> None:
        if not self._persist_path:
            return
        try:
            self._persist_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self._persist_path.with_suffix(".tmp")
            tmp.write_text(json.dumps({"jobs": [_job_to_json(j) for j in self._jobs.values()]}, ensure_ascii=False), encoding="utf-8")
            tmp.replace(self._persist_path)
        except Exception:
            pass

    def _doctor_lock(self, doctor: str) -> threading.Lock:
        with self._lock:
            return self._doctor_locks.setdefault(doctor, threading.Lock())

    @staticmethod
    def _merge(pending: Optional[SaveJob], job: SaveJob) -> SaveJob:
        """Pending months + new months (new entries win, pending base kept).

        A month waiting in the queue from ANOTHER session is a conflict: that
        session's days never reached this editor and would be erased.
        """
        if pending is None:
            return copy.deepcopy(job)
        merged = copy.deepcopy(pending)
        for month, entries in job.entries_by_month.items():
            if month in merged.entries_by_month and merged.origins.get(month) != job.origins.get(month):
                raise ustore.MonthConflictError(f"{int(month[0]):04d}-{int(month[1]):02d}")
            merged.entries_by_month[month] = list(entries)
            merged.base_signatures.setdefault(month, job.base_signatures.get(month, []))
            merged.origins[month] = job.origins.get(month, "")
        merged.doctor_email = job.doctor_email or merged.doctor_email
        merged.by_admin = merged.by_admin or job.by_admin
        merged.next_attempt_at = 0.0
        return merged

    def _continue_from_landed(self, job: SaveJob) -> SaveJob:
        """Same session, month just saved from the queue: its editor base is now
        the landed data (the editor already contains those days)."""
        out = copy.deepcopy(job)
        landed = self._landed.get(job.doctor) or {}
        for month in out.entries_by_month:
            rec = landed.get(month)
            if rec and rec[0] == out.origins.get(month) and ustore.normalize_signature(out.base_signatures.get(month, [])) == rec[1]:
                out.base_signatures[month] = rec[2]
        return out

    def _remember_landed(self, job: SaveJob, result: SaveResult) -> None:
        landed = self._landed.setdefault(job.doctor, {})
        for (yy, mm) in job.entries_by_month:
            landed[(yy, mm)] = (
                job.origins.get((yy, mm), ""),
                ustore.normalize_signature(job.base_signatures.get((yy, mm), [])),
                ustore.month_signature(result.outcome.verified_rows, job.doctor, int(yy), int(mm)),
            )

    def enqueue(self, job: SaveJob) -> None:
        with self._lock:
            self._jobs[job.doctor] = self._merge(self._jobs.get(job.doctor), job)
            self._save()

    def pending(self, doctor: Optional[str] = None) -> list:
        with self._lock:
            jobs = [self._jobs[doctor]] if doctor in self._jobs else ([] if doctor else list(self._jobs.values()))
            return [copy.deepcopy(j) for j in jobs]

    def take_results(self, doctor: str) -> list:
        with self._lock:
            return self._results.pop(doctor, [])

    def _record(self, doctor: str, status: str, message: str, result: Optional[SaveResult] = None) -> None:
        self._results.setdefault(doctor, []).append({
            "status": status,
            "message": message,
            "at": utc_now_iso(),
            "result": result,
        })

    def save_or_enqueue(self, job: SaveJob, *, max_wait: Optional[float] = None):
        """Interactive save: ("saved", SaveResult) or ("queued", None) if GitHub is unavailable.

        Months already waiting in the queue for this doctor are saved together.
        Raises MonthConflictError (nothing written, queue untouched).
        """
        with self._doctor_lock(job.doctor):
            with self._lock:
                merged = self._merge(self._jobs.get(job.doctor), self._continue_from_landed(job))
            try:
                result = save_now(self.cfg, merged, max_wait=max_wait)
            except github_utils.GithubUnavailable as e:
                merged.next_attempt_at = e.retry_at
                merged.last_error = str(e)
                with self._lock:
                    self._jobs[job.doctor] = merged
                    self._save()
                return "queued", None
            with self._lock:
                self._jobs.pop(job.doctor, None)
                self._remember_landed(merged, result)
                self._save()
            return "saved", result

    def process_due(self) -> None:
        now = self._clock()
        with self._lock:
            due = [d for d, j in self._jobs.items() if j.next_attempt_at <= now]
        for doctor in due:
            lock = self._doctor_lock(doctor)
            if not lock.acquire(blocking=False):
                continue  # an interactive save of this doctor is running: it includes the job
            try:
                with self._lock:
                    job = copy.deepcopy(self._jobs.get(doctor))
                if job is not None and job.next_attempt_at <= self._clock():
                    self._process(job)
            finally:
                lock.release()

    def _finish(self, job: SaveJob, status: str, message: str, result: Optional[SaveResult] = None) -> None:
        with self._lock:
            self._jobs.pop(job.doctor, None)
            if result is not None:
                self._remember_landed(job, result)
            self._record(job.doctor, status, message, result)
            self._save()

    def _retry_later(self, job: SaveJob, when: float, error: str) -> None:
        with self._lock:
            current = self._jobs.get(job.doctor)
            if current is not None:
                current.attempts += 1
                current.next_attempt_at = when
                current.last_error = error
            self._save()

    def _process(self, job: SaveJob) -> None:
        try:
            result = save_now(self.cfg, job, max_wait=QUEUE_MAX_WAIT)
        except github_utils.GithubUnavailable as e:
            self._retry_later(job, max(e.retry_at, self._clock() + 5), str(e))
            return
        except ustore.MonthConflictError as e:
            notified = notify_not_saved(
                self.cfg, job,
                f"{e} Per non cancellare quei dati il salvataggio in coda non è stato applicato.",
            )
            self._finish(job, "conflict", f"❌ Salvataggio in coda NON applicato: {e} ({notified})")
            return
        except Exception as e:
            if job.attempts + 1 >= QUEUE_MAX_ATTEMPTS:
                notified = notify_not_saved(self.cfg, job, f"errore ripetuto: {type(e).__name__}: {e}")
                self._finish(job, "failed", f"❌ Salvataggio in coda fallito: {e} ({notified})")
            else:
                self._retry_later(job, self._clock() + min(600, 15 * (2 ** job.attempts)), f"{type(e).__name__}: {e}")
            return

        warnings = write_audit(self.cfg, result, max_wait=QUEUE_MAX_WAIT)
        mail = send_receipt(self.cfg, result)
        if self._on_saved:
            try:
                self._on_saved(result)
            except Exception:
                pass
        message = "✅ Salvataggio in coda completato e verificato sul server."
        if mail:
            message += f" {mail}"
        if warnings:
            message += " " + " ".join(f"⚠️ {w}" for w in warnings)
        self._finish(job, "done", message, result)

    # worker
    def start_worker(self, interval: float = 5.0) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()

        def _loop():
            while not self._stop.is_set():
                try:
                    self.process_due()
                except Exception:
                    pass
                self._stop.wait(interval)

        self._thread = threading.Thread(target=_loop, name="unavailability-save-queue", daemon=True)
        self._thread.start()

    def stop_worker(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=5)
