from __future__ import annotations

import unittest

from tools.corpus.dashboard import _markdown_summary, _part_change_ratio_summary


class DashboardPartStatsTests(unittest.TestCase):
    def test_part_change_ratio_aggregation(self) -> None:
        reports = [
            {
                "display_name": "a.xlsx",
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "part_stats": {
                                "parts_total": 10,
                                "parts_changed": 0,
                                "parts_changed_critical": 0,
                            }
                        },
                    }
                },
            },
            {
                "display_name": "b.xlsx",
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "part_stats": {
                                "parts_total": 20,
                                "parts_changed": 10,
                                "parts_changed_critical": 5,
                            }
                        },
                    }
                },
            },
        ]

        summary = _part_change_ratio_summary(reports)
        self.assertEqual(summary["part_change_ratio"]["count"], 2)
        self.assertAlmostEqual(summary["part_change_ratio"]["mean"], 0.25)
        self.assertAlmostEqual(summary["part_change_ratio"]["p50"], 0.25)
        self.assertAlmostEqual(summary["part_change_ratio"]["p90"], 0.45)
        self.assertAlmostEqual(summary["part_change_ratio"]["max"], 0.5)

        self.assertEqual(summary["part_change_ratio_critical"]["count"], 2)
        self.assertAlmostEqual(summary["part_change_ratio_critical"]["mean"], 0.125)
        self.assertAlmostEqual(summary["part_change_ratio_critical"]["p50"], 0.125)
        self.assertAlmostEqual(summary["part_change_ratio_critical"]["p90"], 0.225)
        self.assertAlmostEqual(summary["part_change_ratio_critical"]["max"], 0.25)

    def test_markdown_includes_part_change_ratio_table(self) -> None:
        reports = [
            {
                "display_name": "a.xlsx",
                "result": {"open_ok": True, "round_trip_ok": True},
            }
        ]
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {
                "total": 1,
                "open_ok": 1,
                "calculate_ok": 0,
                "render_ok": 0,
                "round_trip_ok": 1,
            },
            "rates": {"open": 1.0, "calculate": 0.0, "render": 0.0, "round_trip": 1.0},
            "part_change_ratio": {
                "count": 2,
                "mean": 0.25,
                "p50": 0.25,
                "p90": 0.45,
                "max": 0.5,
            },
            "part_change_ratio_critical": {
                "count": 2,
                "mean": 0.125,
                "p50": 0.125,
                "p90": 0.225,
                "max": 0.25,
            },
        }

        md = _markdown_summary(summary, reports)
        self.assertIn("Part-level change ratio", md)
        self.assertIn("parts_changed / parts_total", md)
        self.assertIn("25.0%", md)
        self.assertIn("45.0%", md)
        self.assertIn("12.5%", md)
        self.assertIn("22.5%", md)


if __name__ == "__main__":
    unittest.main()
