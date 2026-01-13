from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput


def _make_minimal_xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Minimal parts to make this a valid ZIP container for `zipfile.ZipFile` scanning.
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    return buf.getvalue()


class TriageRoundTripFailureKindTests(unittest.TestCase):
    def test_triage_sets_round_trip_failure_kind_from_part_groups(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "parts_with_diffs": [
                                {
                                    "part": "xl/styles.xml",
                                    "group": "styles",
                                    "critical": 1,
                                    "warning": 0,
                                    "info": 0,
                                    "total": 1,
                                }
                            ]
                        },
                    }
                },
                "result": {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "round_trip_fail_on": "critical",
                    "diff_critical_count": 1,
                    "diff_warning_count": 0,
                    "diff_info_count": 0,
                },
            }

            wb = WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx())
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                round_trip_fail_on="critical",
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "round_trip_diff")
        self.assertEqual(report.get("round_trip_failure_kind"), "round_trip_styles")

    def test_triage_respects_round_trip_fail_on_warning(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {
                    "diff": {
                        "status": "ok",
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
                        },
                    }
                },
                "result": {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "round_trip_fail_on": "warning",
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
            }

            wb = WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx())
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                round_trip_fail_on="warning",
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "round_trip_diff")
        self.assertEqual(report.get("round_trip_failure_kind"), "round_trip_styles")

    def test_triage_uses_part_groups_mapping_when_group_missing(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "parts_with_diffs": [
                                {
                                    "part": "docProps/app.xml",
                                    # No `group` field here; should fall back to `part_groups`.
                                    "critical": 0,
                                    "warning": 1,
                                    "info": 0,
                                    "total": 1,
                                }
                            ],
                            "part_groups": {"docProps/app.xml": "doc_props"},
                        },
                    }
                },
                "result": {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "round_trip_fail_on": "warning",
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
            }

            wb = WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx())
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                round_trip_fail_on="warning",
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "round_trip_diff")
        self.assertEqual(report.get("round_trip_failure_kind"), "round_trip_doc_props")


if __name__ == "__main__":
    unittest.main()
