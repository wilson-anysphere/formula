from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.dashboard import _compute_summary
from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput


def _make_minimal_xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # Minimal parts to make this a valid ZIP container for feature scanning.
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    return buf.getvalue()


class TriageFailureCategoryTests(unittest.TestCase):
    def test_triage_sets_calc_mismatch_failure_category(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {
                    "open_ok": True,
                    "round_trip_ok": True,
                    "calculate_ok": False,
                    "render_ok": None,
                },
            }
            report = triage_workbook(
                WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx()),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                round_trip_fail_on="critical",
                recalc=True,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "calc_mismatch")

        summary = _compute_summary([report])
        self.assertEqual(summary.get("failures_by_category", {}).get("calc_mismatch"), 1)

    def test_triage_sets_render_error_failure_category(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {
                    "open_ok": True,
                    "round_trip_ok": True,
                    "calculate_ok": True,
                    "render_ok": False,
                },
            }
            report = triage_workbook(
                WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx()),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                round_trip_fail_on="critical",
                recalc=True,
                render_smoke=True,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "render_error")

        summary = _compute_summary([report])
        self.assertEqual(summary.get("failures_by_category", {}).get("render_error"), 1)


if __name__ == "__main__":
    unittest.main()

