from __future__ import annotations

import json
import io
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path

from tools.corpus.dashboard import _append_trend_file, _trend_entry


class DashboardTrendTests(unittest.TestCase):
    def test_trend_entry_computes_attempted_rates_and_size_overhead(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00+00:00",
            "commit": None,
            "run_url": None,
            "counts": {
                "total": 3,
                "open_ok": 3,
                "round_trip_ok": 2,
                "calculate_ok": 1,
                "calculate_attempted": 2,
                "render_ok": 1,
                "render_attempted": 2,
            },
            "rates": {"open": 1.0, "round_trip": 2 / 3, "calculate": 0.5, "render": 0.5},
            "diff_totals": {"critical": 1, "warning": 2, "info": 3},
            "top_diff_parts_critical": [
                {"part": "xl/workbook.xml", "count": 3},
                {"part": "xl/styles.xml", "count": 1},
            ],
            "top_diff_parts_total": [{"part": "xl/workbook.xml", "count": 4}],
            "top_diff_part_groups_critical": [{"group": "worksheet_xml", "count": 5}],
            "top_diff_part_groups_total": [{"group": "worksheet_xml", "count": 6}],
            "failures_by_category": {"round_trip_diff": 1},
            "failures_by_round_trip_failure_kind": {"round_trip_styles": 1},
            # Size ratios: [110/100=1.1, 180/200=0.9]
            "round_trip_size_overhead": {"count": 2, "mean": 1.0, "p50": 1.0, "p90": 1.08},
            "timings": {
                "load": {"count": 3, "p50_ms": 10, "p90_ms": 20, "mean_ms": 12.0, "max_ms": 30},
                "round_trip": {
                    "count": 3,
                    "p50_ms": 40,
                    "p90_ms": 50,
                    "mean_ms": 41.0,
                    "max_ms": 60,
                },
            },
        }

        entry = _trend_entry(summary)
        self.assertNotIn("commit", entry)
        self.assertNotIn("run_url", entry)

        self.assertEqual(entry["total"], 3)
        self.assertEqual(entry["open_ok"], 3)
        self.assertAlmostEqual(entry["open_rate"], 1.0)
        self.assertAlmostEqual(entry["round_trip_rate"], 2 / 3)

        self.assertEqual(entry["calc_attempted"], 2)
        self.assertEqual(entry["calc_ok"], 1)
        self.assertAlmostEqual(entry["calc_rate"], 0.5)

        self.assertEqual(entry["render_attempted"], 2)
        self.assertEqual(entry["render_ok"], 1)
        self.assertAlmostEqual(entry["render_rate"], 0.5)

        # Size ratios: [1.1, 0.9]
        self.assertEqual(entry["size_overhead_samples"], 2)
        self.assertAlmostEqual(entry["size_overhead_mean"], 1.0)
        self.assertAlmostEqual(entry["size_overhead_p50"], 1.0)
        self.assertAlmostEqual(entry["size_overhead_p90"], 1.08)

        self.assertAlmostEqual(entry["load_p50_ms"], 10.0)
        self.assertAlmostEqual(entry["load_p90_ms"], 20.0)
        self.assertAlmostEqual(entry["round_trip_p50_ms"], 40.0)
        self.assertAlmostEqual(entry["round_trip_p90_ms"], 50.0)

        self.assertEqual(entry["diff_totals"]["critical"], 1)
        self.assertEqual(entry["diff_totals"]["warning"], 2)
        self.assertEqual(entry["diff_totals"]["info"], 3)
        self.assertEqual(entry["diff_totals"]["total"], 6)

        # Optional diff breakdowns should be preserved (top-N only; small in this fixture).
        self.assertEqual(
            entry["top_diff_parts_critical"][0], {"part": "xl/workbook.xml", "count": 3}
        )
        self.assertEqual(
            entry["top_diff_part_groups_total"][0],
            {"group": "worksheet_xml", "count": 6},
        )
        self.assertEqual(
            entry["failures_by_round_trip_failure_kind"], {"round_trip_styles": 1}
        )

    def test_append_trend_file_appends_and_caps(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00+00:00",
            "commit": "abc",
            "run_url": "https://example.invalid/run/1",
            "counts": {"total": 0, "open_ok": 0, "round_trip_ok": 0},
            "rates": {"open": 0.0, "round_trip": 0.0},
        }

        with tempfile.TemporaryDirectory() as td:
            trend_path = Path(td) / "trend.json"
            trend_path.write_text(
                json.dumps(
                    [
                        {"timestamp": "t0", "open_rate": 0.0},
                        {"timestamp": "t1", "open_rate": 0.1},
                        {"timestamp": "t2", "open_rate": 0.2},
                    ]
                ),
                encoding="utf-8",
            )

            entries, prev = _append_trend_file(
                trend_path, summary=summary, max_entries=2
            )

            self.assertEqual(prev, {"timestamp": "t2", "open_rate": 0.2})
            self.assertEqual(len(entries), 2)
            self.assertEqual(entries[0]["timestamp"], "t2")
            self.assertEqual(entries[1]["timestamp"], summary["timestamp"])

            # File should contain the capped list.
            on_disk = json.loads(trend_path.read_text(encoding="utf-8"))
            self.assertEqual(len(on_disk), 2)
            self.assertEqual(on_disk[0]["timestamp"], "t2")

    def test_append_trend_file_treats_whitespace_only_file_as_empty_list(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00+00:00",
            "commit": "abc",
            "run_url": "https://example.invalid/run/1",
            "counts": {"total": 1, "open_ok": 1, "round_trip_ok": 1},
            "rates": {"open": 1.0, "round_trip": 1.0},
        }

        with tempfile.TemporaryDirectory() as td:
            trend_path = Path(td) / "trend.json"
            trend_path.write_text("\n   \n", encoding="utf-8")

            entries, prev = _append_trend_file(trend_path, summary=summary, max_entries=90)

            self.assertIsNone(prev)
            self.assertEqual(len(entries), 1)
            self.assertEqual(entries[0]["timestamp"], summary["timestamp"])

            on_disk = json.loads(trend_path.read_text(encoding="utf-8"))
            self.assertEqual(len(on_disk), 1)

    def test_append_trend_file_overwrites_invalid_json(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00+00:00",
            "commit": "abc",
            "run_url": "https://example.invalid/run/1",
            "counts": {"total": 1, "open_ok": 1, "round_trip_ok": 1},
            "rates": {"open": 1.0, "round_trip": 1.0},
        }

        with tempfile.TemporaryDirectory() as td:
            trend_path = Path(td) / "trend.json"
            trend_path.write_text("{not valid json", encoding="utf-8")

            stderr = io.StringIO()
            with redirect_stderr(stderr):
                entries, prev = _append_trend_file(trend_path, summary=summary, max_entries=90)
            self.assertIn("invalid JSON", stderr.getvalue())

            self.assertIsNone(prev)
            self.assertEqual(len(entries), 1)
            self.assertEqual(entries[0]["timestamp"], summary["timestamp"])

            on_disk = json.loads(trend_path.read_text(encoding="utf-8"))
            self.assertEqual(len(on_disk), 1)


if __name__ == "__main__":
    unittest.main()
