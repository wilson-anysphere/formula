from __future__ import annotations

import unittest

from tools.corpus.trend_delta import trend_delta_markdown


class TrendDeltaTests(unittest.TestCase):
    def test_trend_delta_markdown_renders_expected_sections(self) -> None:
        prev = {
            "timestamp": "t0",
            "open_rate": 0.9,
            "round_trip_rate": 0.8,
            "load_p90_ms": 10,
            "round_trip_p90_ms": 20,
            "calc_rate": 0.5,
            "calc_attempted": 10,
            "render_rate": None,
            "render_attempted": 0,
            "size_overhead_p90": 1.05,
            "size_overhead_samples": 5,
            "part_change_ratio_p90": 0.2,
            "part_change_ratio_critical_p90": 0.1,
            "diff_totals": {"critical": 1, "warning": 2, "info": 3},
            "failures_by_round_trip_failure_kind": {"round_trip_styles": 2},
            "failures_by_category": {"round_trip_diff": 2},
        }
        cur = {
            "timestamp": "t1",
            "open_rate": 1.0,
            "round_trip_rate": 0.9,
            "load_p90_ms": 12,
            "round_trip_p90_ms": 18,
            "calc_rate": 0.6,
            "calc_attempted": 10,
            "render_rate": None,
            "render_attempted": 0,
            "size_overhead_p90": 1.10,
            "size_overhead_samples": 6,
            "part_change_ratio_p90": 0.25,
            "part_change_ratio_critical_p90": 0.15,
            "diff_totals": {"critical": 0, "warning": 1, "info": 3},
            "failures_by_round_trip_failure_kind": {"round_trip_styles": 3},
            "failures_by_category": {"round_trip_diff": 3},
        }

        md = trend_delta_markdown([prev, cur], summary={"timestamp": "t1"})
        self.assertIsNotNone(md)
        assert md is not None
        self.assertIn("## Trend delta", md)
        self.assertIn("Open rate", md)
        self.assertIn("Round-trip rate", md)
        self.assertIn("Load p90", md)
        self.assertIn("Round-trip p90", md)
        self.assertIn("Size ratio p90", md)
        self.assertIn("Part change ratio p90", md)
        self.assertIn("Diff totals", md)
        self.assertIn("Top round-trip failure kinds", md)
        self.assertIn("Top failure categories", md)

    def test_trend_delta_markdown_skips_when_summary_timestamp_mismatch(self) -> None:
        prev = {"timestamp": "t0"}
        cur = {"timestamp": "t1"}
        md = trend_delta_markdown([prev, cur], summary={"timestamp": "t2"})
        self.assertIsNone(md)


if __name__ == "__main__":
    unittest.main()

