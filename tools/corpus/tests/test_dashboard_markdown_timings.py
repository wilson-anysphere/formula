from __future__ import annotations

import unittest

from tools.corpus.dashboard import _markdown_summary


class DashboardMarkdownTimingTests(unittest.TestCase):
    def test_markdown_includes_timings_section(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {
                "total": 2,
                "open_ok": 2,
                "calculate_ok": 0,
                "calculate_attempted": 0,
                "render_ok": 0,
                "render_attempted": 0,
                "round_trip_ok": 2,
            },
            "rates": {"open": 1.0, "calculate": None, "render": None, "round_trip": 1.0},
            "diff_totals": {},
            "round_trip_size_overhead": {"count": 0},
            "timings": {
                "load": {
                    "count": 2,
                    "mean_ms": 150.0,
                    "p50_ms": 150.0,
                    "p90_ms": 200.0,
                    "max_ms": 200,
                },
                "round_trip": {
                    "count": 2,
                    "mean_ms": 250.0,
                    "p50_ms": 250.0,
                    "p90_ms": 300.0,
                    "max_ms": 300,
                },
                "diff": {"count": 0, "mean_ms": None, "p50_ms": None, "p90_ms": None, "max_ms": None},
                "recalc": {"count": 0, "mean_ms": None, "p50_ms": None, "p90_ms": None, "max_ms": None},
                "render": {"count": 0, "mean_ms": None, "p50_ms": None, "p90_ms": None, "max_ms": None},
            },
        }

        md = _markdown_summary(summary, reports=[])
        self.assertIn("## Timings", md)
        self.assertIn("| Step | Count | Mean (ms) | P50 (ms) | P90 (ms) | Max (ms) |", md)
        self.assertIn("| load | 2 | 150 | 150 | 200 | 200 |", md)
        self.assertIn("| round_trip | 2 | 250 | 250 | 300 | 300 |", md)

        # Timings should appear before the size overhead section for readability.
        self.assertLess(md.find("## Timings"), md.find("## Round-trip size overhead"))


if __name__ == "__main__":
    unittest.main()

