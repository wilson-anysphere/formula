from __future__ import annotations

import unittest

from tools.corpus.dashboard import _markdown_summary, _round_trip_size_overhead


class DashboardOverheadTests(unittest.TestCase):
    def test_round_trip_size_overhead_stats(self) -> None:
        reports = [
            {
                "display_name": "a.xlsx",
                "size_bytes": 100,
                "steps": {
                    "round_trip": {
                        "status": "ok",
                        "details": {"output_size_bytes": 103},
                    }
                },
            },
            {
                "display_name": "b.xlsx",
                "size_bytes": 100,
                "steps": {
                    "round_trip": {
                        "status": "ok",
                        "details": {"output_size_bytes": 110},
                    }
                },
            },
            {
                "display_name": "c.xlsx",
                "size_bytes": 100,
                "steps": {
                    "round_trip": {
                        "status": "ok",
                        "details": {"output_size_bytes": 120},
                    }
                },
            },
            # Missing / non-ok round-trip steps should be ignored.
            {"display_name": "missing.xlsx", "size_bytes": 100, "steps": {}},
            {
                "display_name": "skipped.xlsx",
                "size_bytes": 100,
                "steps": {"round_trip": {"status": "skipped", "details": {"reason": "x"}}},
            },
            {
                "display_name": "failed.xlsx",
                "size_bytes": 100,
                "steps": {"round_trip": {"status": "failed"}},
            },
        ]

        stats = _round_trip_size_overhead(reports)
        self.assertEqual(stats["count"], 3)
        self.assertAlmostEqual(stats["mean"], 1.11, places=6)
        self.assertAlmostEqual(stats["p50"], 1.10, places=6)
        self.assertAlmostEqual(stats["p90"], 1.18, places=6)
        self.assertAlmostEqual(stats["max"], 1.20, places=6)
        self.assertEqual(stats["count_over_1_05"], 2)
        self.assertEqual(stats["count_over_1_10"], 1)

    def test_dashboard_markdown_includes_round_trip_overhead_section(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {
                "total": 3,
                "open_ok": 3,
                "calculate_ok": 0,
                "render_ok": 0,
                "round_trip_ok": 3,
            },
            "rates": {"open": 1.0, "calculate": 0.0, "render": 0.0, "round_trip": 1.0},
            "round_trip_size_overhead": {
                "count": 3,
                "mean": 1.11,
                "p50": 1.10,
                "p90": 1.18,
                "max": 1.20,
                "count_over_1_05": 2,
                "count_over_1_10": 1,
            },
        }
        reports = [
            {"display_name": "a.xlsx", "result": {"open_ok": True, "round_trip_ok": True}},
            {"display_name": "b.xlsx", "result": {"open_ok": True, "round_trip_ok": True}},
            {"display_name": "c.xlsx", "result": {"open_ok": True, "round_trip_ok": True}},
        ]

        md = _markdown_summary(summary, reports)
        self.assertIn("## Round-trip size overhead", md)
        self.assertIn("Workbooks with size data: **3**", md)
        self.assertIn("mean **1.110**", md)
        self.assertIn("p50 **1.100**", md)
        self.assertIn("p90 **1.180**", md)
        self.assertIn("max **1.200**", md)
        self.assertIn("Exceeding ratio thresholds (>1.05 / >1.10): **2 / 1**", md)


if __name__ == "__main__":
    unittest.main()

