from __future__ import annotations

import unittest

from tools.corpus.dashboard import _markdown_summary
from tools.corpus.triage import _compare_expectations


class TriageSchemaTests(unittest.TestCase):
    def test_compare_expectations_supports_numeric_thresholds(self) -> None:
        reports = [
            {
                "display_name": "book.xlsx",
                "result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0},
            }
        ]
        expectations = {
            "book.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}
        }
        regressions, improvements = _compare_expectations(reports, expectations)
        self.assertEqual(regressions, [])
        self.assertEqual(improvements, [])

        # Numeric regressions: higher-than-expected counts should fail CI.
        reports[0]["result"]["diff_critical_count"] = 2
        regressions, _ = _compare_expectations(reports, expectations)
        self.assertTrue(any("diff_critical_count" in r for r in regressions))

        # Numeric improvements: lower-than-expected counts are surfaced as improvements.
        expectations["book.xlsx"]["diff_critical_count"] = 2
        reports[0]["result"]["diff_critical_count"] = 0
        _, improvements = _compare_expectations(reports, expectations)
        self.assertTrue(any("diff_critical_count" in r for r in improvements))

    def test_compare_expectations_treats_skips_as_regressions(self) -> None:
        reports = [{"display_name": "book.xlsx", "result": {"open_ok": None}}]
        expectations = {"book.xlsx": {"open_ok": True}}
        regressions, _ = _compare_expectations(reports, expectations)
        self.assertEqual(len(regressions), 1)

    def test_dashboard_markdown_includes_diff_and_render_columns(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {
                "total": 1,
                "open_ok": 1,
                "calculate_ok": 0,
                "calculate_attempted": 0,
                "render_ok": 0,
                "render_attempted": 0,
                "round_trip_ok": 1,
            },
            "rates": {"open": 1.0, "calculate": None, "render": None, "round_trip": 1.0},
            "diff_totals": {"critical": 0, "warning": 1, "info": 0},
        }
        reports = [
            {
                "display_name": "book.xlsx",
                "result": {
                    "open_ok": True,
                    "calculate_ok": None,
                    "render_ok": None,
                    "round_trip_ok": True,
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
            }
        ]
        md = _markdown_summary(summary, reports)
        self.assertIn("Diff (C/W/I)", md)
        self.assertIn("0/1/0", md)
        # Calculate/render should not be reported as "0.0%" when triage skipped those steps.
        self.assertIn("Calculate: **0 / 0** (SKIP", md)
        self.assertIn("Render: **0 / 0** (SKIP", md)


if __name__ == "__main__":
    unittest.main()
