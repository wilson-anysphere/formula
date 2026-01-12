from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput


def _make_xlsx_with_cell_images() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellimages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ci:cellImages xmlns:ci="http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <ci:cellImage><a:blip r:embed="rId1"/></ci:cellImage>
  <ci:cellImage><a:blip r:embed="rId2"/></ci:cellImage>
</ci:cellImages>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99"
                Type="http://example.com/relationships/cellImages"
                Target="cellimages.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://example.com/relationships/image" Target="media/image2.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_rid_on_cellimage() -> bytes:
    """Variant fixture where the relationship id is stored as `r:id` on `<cellImage>`."""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellimages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>
""",
        )
        z.writestr(
            "xl/_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_numeric_suffix() -> bytes:
    """Synthetic fixture using `xl/cellimages1.xml` (numeric suffix)."""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellimages1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ci:cellImages xmlns:ci="http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <ci:cellImage><a:blip r:embed="rId1"/></ci:cellImage>
</ci:cellImages>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99"
                Type="http://example.com/relationships/cellImages"
                Target="cellimages1.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/_rels/cellimages1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


class TriageCellImagesTests(unittest.TestCase):
    def test_triage_extracts_cell_images_metadata(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(display_name="book.xlsx", data=_make_xlsx_with_cell_images())
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

        self.assertTrue(report["features"]["has_cell_images"])
        self.assertIn("cell_images", report)

        cell_images = report["cell_images"]
        self.assertEqual(cell_images["part_name"], "xl/cellimages.xml")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml")
        self.assertEqual(cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages")
        self.assertEqual(cell_images["root_local_name"], "cellImages")
        self.assertEqual(
            cell_images["root_namespace"],
            "http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages",
        )
        self.assertEqual(cell_images["embed_rids_count"], 2)
        self.assertEqual(
            cell_images["rels_types"],
            [
                "http://example.com/relationships/image",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            ],
        )

    def test_triage_extracts_cell_images_metadata_numeric_suffix(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_numeric_suffix()
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

        self.assertTrue(report["features"]["has_cell_images"])
        self.assertIn("cell_images", report)

        cell_images = report["cell_images"]
        self.assertEqual(cell_images["part_name"], "xl/cellimages1.xml")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml")
        self.assertEqual(cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages")
        self.assertEqual(cell_images["root_local_name"], "cellImages")
        self.assertEqual(
            cell_images["root_namespace"],
            "http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages",
        )
        self.assertEqual(cell_images["embed_rids_count"], 1)
        self.assertEqual(
            cell_images["rels_types"],
            ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"],
        )

    def test_triage_counts_embed_rids_outside_blips(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_rid_on_cellimage()
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

        self.assertTrue(report["features"]["has_cell_images"])
        self.assertIn("cell_images", report)

        # Previously we only counted `<a:blip r:embed="...">`. This fixture uses
        # `<cellImage r:id="...">` and should still be counted.
        self.assertEqual(report["cell_images"]["embed_rids_count"], 1)


if __name__ == "__main__":
    unittest.main()
