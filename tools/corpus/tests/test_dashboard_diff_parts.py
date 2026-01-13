from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

from tools.corpus import dashboard as dashboard_mod


class DashboardDiffPartsTests(unittest.TestCase):
    def test_dashboard_reports_top_diff_parts_when_present(self) -> None:
        with tempfile.TemporaryDirectory(prefix="corpus-dashboard-") as tmp:
            triage_dir = Path(tmp)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            report = {
                "display_name": "book.xlsx",
                "result": {
                    "open_ok": True,
                    "calculate_ok": None,
                    "render_ok": None,
                    "round_trip_ok": False,
                    "diff_critical_count": 3,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
                "steps": {
                    "diff": {
                        "status": "ok",
                        "duration_ms": 1,
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 3, "warning": 1, "info": 0, "total": 4},
                            "equal": False,
                            "parts_with_diffs": [
                                {
                                    "part": "xl/workbook.xml",
                                    "group": "other",
                                    "critical": 2,
                                    "warning": 0,
                                    "info": 0,
                                    "total": 2,
                                },
                                {
                                    "part": "docProps/app.xml",
                                    "critical": 1,
                                    "warning": 0,
                                    "info": 0,
                                    "total": 1,
                                },
                                {
                                    "part": "xl/theme/theme1.xml",
                                    "group": "other",
                                    "critical": 0,
                                    "warning": 1,
                                    "info": 0,
                                    "total": 1,
                                },
                            ],
                            "critical_parts": ["xl/workbook.xml", "docProps/app.xml"],
                            "part_groups": {
                                "docProps/app.xml": "doc_props",
                                "xl/theme/theme1.xml": "other",
                                "xl/workbook.xml": "other",
                            },
                            "top_differences": [],
                        },
                    }
                },
            }

            (reports_dir / "r1.json").write_text(
                json.dumps(report, indent=2, sort_keys=True), encoding="utf-8"
            )

            old_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                ]
                rc = dashboard_mod.main()
            finally:
                sys.argv = old_argv

            self.assertEqual(rc, 0)

            md = (triage_dir / "summary.md").read_text(encoding="utf-8")
            self.assertIn("## Top diff parts (CRITICAL)", md)
            self.assertIn("| xl/workbook.xml | 2 |", md)
            self.assertIn("| docProps/app.xml | 1 |", md)
            self.assertIn("## Top diff parts (all severities)", md)
            self.assertIn("| xl/workbook.xml | 2 |", md)

            self.assertIn("## Top diff part groups (CRITICAL)", md)
            self.assertIn("| other | 2 |", md)
            self.assertIn("| doc_props | 1 |", md)

            self.assertIn("## Top diff part groups (all severities)", md)
            self.assertIn("| other | 3 |", md)
            self.assertIn("| doc_props | 1 |", md)


if __name__ == "__main__":
    unittest.main()
