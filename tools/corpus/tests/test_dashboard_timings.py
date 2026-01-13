from __future__ import annotations

import unittest

from tools.corpus.dashboard import _compute_timings


class DashboardTimingTests(unittest.TestCase):
    def test_compute_timings_median_and_p90(self) -> None:
        # Ten samples so the p90 is not equal to the max (interpolated percentile).
        load_durations = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]
        reports = [
            {
                "display_name": f"book-{d}.xlsx",
                "steps": {
                    "load": {"status": "ok", "duration_ms": d},
                    "round_trip": {"status": "ok", "duration_ms": d + 1},
                },
            }
            for d in load_durations
        ]
        # Failed steps should not contribute to timing percentiles.
        reports.append(
            {
                "display_name": "failed.xlsx",
                "steps": {"load": {"status": "failed", "duration_ms": 5000}},
            }
        )
        # Include a report without steps to ensure it's ignored.
        reports.append({"display_name": "nosteps.xlsx"})

        timings = _compute_timings(reports)

        load = timings["load"]
        self.assertEqual(load["count"], 10)
        self.assertEqual(load["mean_ms"], 550.0)
        self.assertEqual(load["p50_ms"], 550.0)
        self.assertEqual(load["p90_ms"], 910.0)
        self.assertEqual(load["max_ms"], 1000)

        round_trip = timings["round_trip"]
        self.assertEqual(round_trip["count"], 10)
        self.assertEqual(round_trip["p50_ms"], 551.0)
        self.assertEqual(round_trip["p90_ms"], 911.0)

        # Optional steps (recalc/render) are skipped by default and should have no samples.
        self.assertEqual(timings["recalc"]["count"], 0)
        self.assertIsNone(timings["recalc"]["p50_ms"])
        self.assertEqual(timings["render"]["count"], 0)
        self.assertIsNone(timings["render"]["p90_ms"])


if __name__ == "__main__":
    unittest.main()
