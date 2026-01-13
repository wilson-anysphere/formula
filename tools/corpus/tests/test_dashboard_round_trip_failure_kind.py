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
                            # Match the Rust helper schema: list of per-part diff summaries.
                            "parts_with_diffs": [
                                {"part": "_rels/.rels", "group": "rels", "critical": 1},
                                {
                                    "part": "xl/_rels/workbook.xml.rels",
                                    "group": "rels",
                                    "critical": 1,
                                },
                            ],
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
                            "parts_with_diffs": [
                                {
                                    "part": "[Content_Types].xml",
                                    "group": "content_types",
                                    "critical": 1,
                                },
                                {
                                    "part": "xl/_rels/workbook.xml.rels",
                                    "group": "rels",
                                    "critical": 1,
                                },
                            ],
                        }
                    }
                },
            },
            {
                "display_name": "styles.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": [
                                {"part": "xl/styles.xml", "group": "styles", "critical": 1}
                            ]
                        }
                    }
                },
            },
            {
                "display_name": "worksheets.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 2},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": [
                                {
                                    "part": "xl/worksheets/sheet1.xml",
                                    "group": "worksheet_xml",
                                    "critical": 1,
                                },
                                {
                                    "part": "xl/worksheets/sheet2.xml",
                                    "group": "worksheet_xml",
                                    "critical": 1,
                                },
                            ],
                        }
                    }
                },
            },
            {
                "display_name": "other.xlsx",
                "failure_category": "round_trip_diff",
                "result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": [
                                {"part": "xl/workbook.xml", "group": "other", "critical": 1}
                            ]
                        }
                    }
                },
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

    def test_dashboard_infers_kind_for_warning_fail_on(self) -> None:
        # If round_trip_ok is false due to a WARN diff (fail-on=warning), we should still bucket
        # based on the WARN-changing parts/groups rather than falling back to round_trip_other.
        reports = [
            {
                "display_name": "warn_styles.xlsx",
                "failure_category": "round_trip_diff",
                "result": {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "round_trip_fail_on": "warning",
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
                "steps": {
                    "diff": {
                        "details": {
                            "parts_with_diffs": [
                                {
                                    "part": "xl/styles.xml",
                                    "group": "styles",
                                    "critical": 0,
                                    "warning": 1,
                                    "info": 0,
                                    "total": 1,
                                }
                            ]
                        }
                    }
                },
            }
        ]

        summary = _compute_summary(reports)
        self.assertEqual(summary["failures_by_round_trip_failure_kind"], {"round_trip_styles": 1})


if __name__ == "__main__":
    unittest.main()
