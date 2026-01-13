from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.ingest import _triage_sanitized_workbook
from tools.corpus.util import WorkbookInput


def _make_minimal_xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</workbook>
""",
        )
    return buf.getvalue()


class IngestTriageInvocationTests(unittest.TestCase):
    def test_triage_wrapper_uses_triage_defaults_without_rust_build(self) -> None:
        import tools.corpus.triage as triage_mod

        observed: dict[str, object] = {}

        original_build_rust_helper = triage_mod._build_rust_helper
        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_run_rust_triage(
                exe: Path,
                workbook_bytes: bytes,
                *,
                workbook_name: str,
                diff_ignore: set[str],
                diff_ignore_path: tuple[str, ...] | list[str] = (),
                diff_ignore_path_in: tuple[str, ...] | list[str] = (),
                diff_limit: int,
                recalc: bool,
                render_smoke: bool,
                strict_calc_chain: bool = False,
            ) -> dict:
                observed["exe"] = exe
                observed["workbook_name"] = workbook_name
                observed["diff_ignore"] = diff_ignore
                observed["diff_ignore_path"] = tuple(diff_ignore_path)
                observed["diff_ignore_path_in"] = tuple(diff_ignore_path_in)
                observed["diff_limit"] = diff_limit
                observed["recalc"] = recalc
                observed["render_smoke"] = render_smoke
                observed["strict_calc_chain"] = strict_calc_chain
                return {"steps": {}, "result": {"open_ok": True, "round_trip_ok": True}}

            triage_mod._run_rust_triage = _fake_run_rust_triage  # type: ignore[assignment]

            report = _triage_sanitized_workbook(
                WorkbookInput(display_name="book.xlsx", data=_make_minimal_xlsx())
            )
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(observed["exe"], Path("noop"))
        self.assertEqual(observed["workbook_name"], "book.xlsx")
        self.assertEqual(observed["diff_limit"], 25)
        self.assertEqual(observed["recalc"], False)
        self.assertEqual(observed["render_smoke"], False)
        self.assertEqual(observed["workbook_name"], "book.xlsx")

        diff_ignore = observed["diff_ignore"]
        self.assertIsInstance(diff_ignore, set)
        self.assertTrue(triage_mod.DEFAULT_DIFF_IGNORE.issubset(diff_ignore))

        self.assertTrue(report["result"]["open_ok"])
        self.assertTrue(report["result"]["round_trip_ok"])


if __name__ == "__main__":
    unittest.main()
