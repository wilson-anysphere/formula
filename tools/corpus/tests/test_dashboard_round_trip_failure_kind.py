from __future__ import annotations

import unittest

from tools.corpus.dashboard import _compute_summary


class DashboardRoundTripFailureKindTests(unittest.TestCase):
    def test_dashboard_groups_round_trip_failures_by_kind(self) -> None:
        reports = [
            {
                "display_name": "rels.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 2},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": {
                                "critical": [
                                    "_rels/.rels",
                                    "xl/_rels/workbook.xml.rels",
                                ]
                            }
                        }
                    }
                },
            },
            {
                "display_name": "content_types.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": {
                                "critical": [
                                    "[Content_Types].xml",
                                    "xl/_rels/workbook.xml.rels",
                                ]
                            }
                        }
                    }
                },
            },
            {
                "display_name": "styles.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                "steps": {"diff": {"details": {"parts_with_diffs": {"critical": ["xl/styles.xml"]}}}},
            },
            {
                "display_name": "worksheets.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 2},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": {
                                "critical": [
                                    "xl/worksheets/sheet1.xml",
                                    "xl/worksheets/sheet2.xml",
                                ]
                            }
                        }
                    }
                },
            },
            {
                "display_name": "other.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                "steps": {"diff": {"details": {"parts_with_diffs": {"critical": ["xl/workbook.xml"]}}}},
            },
            # Non-round-trip failure categories should not be counted in the round-trip buckets.
            {
                "display_name": "open_error.xlsx",
                "failure_category": "open_error",
                "result": {"open_ok": False, "round_trip_ok": False},
            },
            {
                "display_name": "pass.xlsx",
                "result": {"open_ok": True, "round_trip_ok": True},
            },
        ]

        summary = _compute_summary(reports)
        self.assertEqual(
            summary["failures_by_round_trip_failure_kind"],
            {
                "round_trip_content_types": 1,
                "round_trip_other": 1,
                "round_trip_rels": 1,
                "round_trip_styles": 1,
                "round_trip_worksheets": 1,
            },
        )


if __name__ == "__main__":
    unittest.main()

