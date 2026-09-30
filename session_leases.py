# -*- coding: utf-8 -*-
"""Single active session per doctor, held in the app process.

All Streamlit sessions of the app run in one Python process, so the "who is
editing" lease lives in memory: no GitHub commit per login/heartbeat (they were
half of all commits and a read every 5 seconds per active doctor).
If the process restarts, leases are simply free again.
"""

from __future__ import annotations

import threading
import time
from typing import Callable, Optional


class LeaseRegistry:
    def __init__(self, ttl_seconds: float = 1200, clock: Callable[[], float] = time.time):
        self._ttl = ttl_seconds
        self._clock = clock
        self._lock = threading.Lock()
        self._leases: dict[str, tuple[str, float]] = {}

    def _current(self, doctor: str) -> Optional[str]:
        lease = self._leases.get(doctor)
        if lease is None or self._clock() - lease[1] > self._ttl:
            return None
        return lease[0]

    def owner(self, doctor: str) -> Optional[str]:
        with self._lock:
            return self._current(doctor)

    def acquire(self, doctor: str, session_id: str) -> None:
        """New login: this session takes over (older sessions get kicked out)."""
        with self._lock:
            self._leases[doctor] = (session_id, self._clock())

    def is_active(self, doctor: str, session_id: str) -> bool:
        """True if this session owns the lease, or nobody holds a live one."""
        with self._lock:
            current = self._current(doctor)
            return current is None or current == session_id

    def touch(self, doctor: str, session_id: str) -> bool:
        """Heartbeat: refresh our lease, or reclaim it if free. Never steal a live one."""
        with self._lock:
            current = self._current(doctor)
            if current is not None and current != session_id:
                return False
            self._leases[doctor] = (session_id, self._clock())
            return True

    def release(self, doctor: str, session_id: str) -> None:
        with self._lock:
            if self._current(doctor) == session_id:
                self._leases.pop(doctor, None)
