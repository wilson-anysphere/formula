from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput


def _make_xlsx_with_styles_xml(styles_xml: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/styles.xml", styles_xml)
    return buf.getvalue()


class TriageStyleStatsTests(unittest.TestCase):
    def test_triage_extracts_style_counts(self) -> None:
        import tools.corpus.triage as triage_mod

        styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.00"/>
    <numFmt numFmtId="165" formatCode="0.0"/>
  </numFmts>
  <fonts count="3"><font/><font/><font/></fonts>
  <fills><fill/><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="7"><xf/><xf/><xf/><xf/><xf/><xf/><xf/></cellXfs>
  <cellStyles><cellStyle/><cellStyle/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="9" defaultTableStyle="TableStyleMedium9"/>
  <extLst><ext uri="{00000000-0000-0000-0000-000000000000}"/><ext uri="{11111111-1111-1111-1111-111111111111}"/></extLst>
</styleSheet>
"""

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_styles_xml(styles_xml)
            )
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertIn("style_stats", report)
        stats = report["style_stats"]
        self.assertEqual(stats["numFmts"], 2)
        self.assertEqual(stats["fonts"], 3)
        self.assertEqual(stats["fills"], 2)
        self.assertEqual(stats["borders"], 1)
        self.assertEqual(stats["cellStyleXfs"], 1)
        self.assertEqual(stats["cellXfs"], 7)
        self.assertEqual(stats["cellStyles"], 2)
        self.assertEqual(stats["dxfs"], 0)
        self.assertEqual(stats["tableStyles"], 9)
        self.assertEqual(stats["extLst"], 2)

        for k, v in stats.items():
            self.assertIsInstance(v, int, msg=f"{k} should be int")

    def test_triage_tolerates_malformed_styles_xml(self) -> None:
        import tools.corpus.triage as triage_mod

        styles_xml = "<styleSheet><fonts></styleSheet>"  # malformed XML

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_styles_xml(styles_xml)
            )
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertNotIn("style_stats", report)
        self.assertIn("style_stats_error", report)
        self.assertTrue(str(report["style_stats_error"]))

    def test_private_mode_hashes_style_stats_error(self) -> None:
        import tools.corpus.triage as triage_mod

        styles_xml = "<styleSheet><fonts></styleSheet>"  # malformed XML

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_styles_xml(styles_xml)
            )
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
                privacy_mode="private",
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertNotIn("style_stats", report)
        self.assertIn("style_stats_error", report)
        self.assertTrue(str(report["style_stats_error"]).startswith("sha256="))


if __name__ == "__main__":
    unittest.main()
