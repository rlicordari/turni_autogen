import threading
import unittest

from session_leases import LeaseRegistry


class Clock:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t


class LeaseRegistryTests(unittest.TestCase):
    def setUp(self):
        self.clock = Clock()
        self.reg = LeaseRegistry(ttl_seconds=1200, clock=self.clock)

    def test_new_login_kicks_out_previous_session(self):
        self.reg.acquire("Migliorato", "pc")
        self.reg.acquire("Migliorato", "phone")

        self.assertFalse(self.reg.is_active("Migliorato", "pc"))
        self.assertTrue(self.reg.is_active("Migliorato", "phone"))

    def test_heartbeat_of_old_session_does_not_steal_the_lease(self):
        self.reg.acquire("Migliorato", "pc")
        self.reg.acquire("Migliorato", "phone")

        self.assertFalse(self.reg.touch("Migliorato", "pc"))
        self.assertTrue(self.reg.is_active("Migliorato", "phone"))

    def test_expired_lease_can_be_reclaimed_by_heartbeat(self):
        self.reg.acquire("Migliorato", "phone")
        self.clock.t += 1201

        self.assertTrue(self.reg.is_active("Migliorato", "pc"))
        self.assertTrue(self.reg.touch("Migliorato", "pc"))
        self.assertFalse(self.reg.is_active("Migliorato", "phone"))

    def test_release_frees_only_own_lease(self):
        self.reg.acquire("Migliorato", "phone")
        self.reg.release("Migliorato", "pc")
        self.assertFalse(self.reg.is_active("Migliorato", "pc"))

        self.reg.release("Migliorato", "phone")
        self.assertTrue(self.reg.is_active("Migliorato", "pc"))

    def test_doctors_are_independent(self):
        self.reg.acquire("Migliorato", "a")
        self.reg.acquire("Crea", "b")
        self.assertTrue(self.reg.is_active("Migliorato", "a"))
        self.assertTrue(self.reg.is_active("Crea", "b"))

    def test_concurrent_logins_leave_exactly_one_owner(self):
        barrier = threading.Barrier(16)

        def login(sid):
            barrier.wait()
            self.reg.acquire("Migliorato", sid)

        threads = [threading.Thread(target=login, args=(f"s{i}",)) for i in range(16)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        owners = [f"s{i}" for i in range(16) if self.reg.owner("Migliorato") == f"s{i}"]
        self.assertEqual(len(owners), 1)


if __name__ == "__main__":
    unittest.main()
