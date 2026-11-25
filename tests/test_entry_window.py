import unittest
from datetime import datetime, timezone

import entry_window as ew

R = ew.LOCAL_TZ


class EntryWindowTests(unittest.TestCase):
    def test_next_month_closes_on_the_21st_at_midnight_italian_time(self):
        self.assertTrue(ew.month_is_open(2026, 12, datetime(2026, 11, 20, 23, 59, tzinfo=R), 21))
        self.assertFalse(ew.month_is_open(2026, 12, datetime(2026, 11, 21, 0, 0, tzinfo=R), 21))
        # Server in UTC: 20/11 23:30 UTC = 21/11 00:30 a Roma → già chiuso
        self.assertFalse(ew.month_is_open(2026, 12, datetime(2026, 11, 20, 23, 30, tzinfo=timezone.utc), 21))

    def test_later_months_stay_open_and_january_uses_december(self):
        now = datetime(2026, 11, 25, 10, 0, tzinfo=R)
        self.assertTrue(ew.month_is_open(2027, 1, now, 21))
        self.assertFalse(ew.month_is_open(2026, 11, now, 21))   # mese in corso
        self.assertEqual(ew.month_deadline(2027, 1, 21), datetime(2026, 12, 21, tzinfo=R))

    def test_disabled_and_out_of_range_days(self):
        now = datetime(2026, 11, 25, tzinfo=R)
        self.assertTrue(ew.month_is_open(2026, 12, now, 0))
        self.assertIsNone(ew.month_deadline(2026, 12, 0))
        self.assertEqual(ew.normalize_close_day(31), 28)
        self.assertEqual(ew.normalize_close_day("x"), ew.DEFAULT_CLOSE_DAY)

    def test_first_open_month_skips_the_closed_one(self):
        self.assertEqual(ew.first_open_month(datetime(2026, 11, 10, tzinfo=R), 21), (2026, 12))
        self.assertEqual(ew.first_open_month(datetime(2026, 11, 21, 8, tzinfo=R), 21), (2027, 1))
        self.assertEqual(ew.first_open_month(datetime(2026, 12, 22, tzinfo=R), 21), (2027, 2))


if __name__ == "__main__":
    unittest.main()
