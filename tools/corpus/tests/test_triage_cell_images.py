from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput


def _rewrite_zip_with_leading_slash_entry_names(data: bytes) -> bytes:
    zin_buf = io.BytesIO(data)
    zout_buf = io.BytesIO()
    with zipfile.ZipFile(zin_buf, "r") as zin:
        with zipfile.ZipFile(zout_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.is_dir():
                    continue
                name = info.filename
                new_name = name if name.startswith("/") else f"/{name}"
                zout.writestr(new_name, zin.read(name))
    return zout_buf.getvalue()


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


def _make_xlsx_with_cell_images_malformed_paths() -> bytes:
    """Fixture with malformed OPC paths that require `..` resolution.

    - `[Content_Types].xml` Override `PartName` includes backslashes + `..`
    - The `cellimages.xml.rels` ZIP entry name includes `..` segments
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl\\_rels\\..\\cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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
        # Intentionally store the .rels part under a malformed name that normalizes to the
        # expected `xl/_rels/cellimages.xml.rels`.
        z.writestr(
            "xl/_rels/../_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://example.com/relationships/image" Target="media/image2.png"/>
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


def _make_xlsx_with_cell_images_multiple_parts() -> bytes:
    """Fixture containing both `xl/cellimages.xml` and `xl/cellimages1.xml`.

    Triage should prefer `xl/cellimages.xml` when present.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
  <Override PartName="/xl/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml+1"/>
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
            "xl/cellimages1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ci:cellImages xmlns:ci="http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <ci:cellImage><a:blip r:embed="rId9"/></ci:cellImage>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_smallest_numeric_suffix() -> bytes:
    """Fixture with `cellimages2.xml` and `cellimages10.xml` (no `cellimages.xml`).

    Triage should select the smallest numeric suffix (2).
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages2.xml" ContentType="application/vnd.ms-excel.cellimages+xml+2"/>
  <Override PartName="/xl/cellimages10.xml" ContentType="application/vnd.ms-excel.cellimages+xml+10"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellimages2.xml",
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
            "xl/cellimages10.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ci:cellImages xmlns:ci="http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <ci:cellImage><a:blip r:embed="rId9"/></ci:cellImage>
</ci:cellImages>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99"
                Type="http://example.com/relationships/cellImages2"
                Target="xl/cellimages2.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/_rels/cellimages2.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_folder_layout_and_basename_target() -> bytes:
    """Folder layout where `workbook.xml.rels` targets only the basename.

    This is technically non-standard, but triage should still extract the relationship type.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml+folder"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellimages/cellimages1.xml",
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
                Type="http://example.com/relationships/cellImages-folder"
                Target="cellimages1.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/cellimages/_rels/cellimages1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_uppercase_part_name() -> bytes:
    """Fixture where the cellImages part name casing differs (`CellImages1.XML`)."""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/CellImages1.XML" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/CellImages1.XML",
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
                Target="CellImages1.XML"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/_rels/CellImages1.XML.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


def _make_xlsx_with_cell_images_duplicate_basename_workbook_targets_root() -> bytes:
    """Fixture with two parts that share the same basename.

    - `xl/cellimages1.xml` (root)
    - `xl/cellimages/cellimages1.xml` (folder layout)

    `workbook.xml.rels` targets `cellimages1.xml`, which resolves to `xl/cellimages1.xml`
    and exists in the package. Triage should prefer the root part when multiple candidates
    share the same numeric suffix.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml+root"/>
  <Override PartName="/xl/cellimages/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml+folder"/>
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
            "xl/cellimages/cellimages1.xml",
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
                Type="http://example.com/relationships/cellImages-root"
                Target="cellimages1.xml"/>
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

    def test_triage_tolerates_leading_slash_zip_entry_names(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx",
                data=_rewrite_zip_with_leading_slash_entry_names(_make_xlsx_with_cell_images()),
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
        self.assertEqual(report["cell_images"]["part_name"], "/xl/cellimages.xml")

    def test_triage_normalizes_malformed_part_names_and_rels_paths(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_malformed_paths()
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
        self.assertEqual(cell_images["part_name"], "xl/cellimages.xml")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml")
        self.assertEqual(
            cell_images["rels_types"],
            [
                "http://example.com/relationships/image",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            ],
        )

    def test_triage_prefers_cellimages_xml_when_present(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_multiple_parts()
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
        self.assertEqual(cell_images["part_name"], "xl/cellimages.xml")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml")
        self.assertEqual(cell_images["embed_rids_count"], 2)

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

    def test_triage_chooses_smallest_numeric_suffix(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_smallest_numeric_suffix()
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
        self.assertEqual(cell_images["part_name"], "xl/cellimages2.xml")
        # Relationship targets normally resolve relative to `xl/workbook.xml`, but be tolerant of
        # producers that include an `xl/` prefix.
        self.assertEqual(cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages2")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml+2")
        self.assertEqual(cell_images["embed_rids_count"], 2)

    def test_triage_matches_workbook_rel_by_basename_for_folder_layout(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx",
                data=_make_xlsx_with_cell_images_folder_layout_and_basename_target(),
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
        self.assertEqual(cell_images["part_name"], "xl/cellimages/cellimages1.xml")
        self.assertEqual(cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages-folder")

    def test_triage_extracts_cell_images_metadata_case_insensitively(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx", data=_make_xlsx_with_cell_images_uppercase_part_name()
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
        self.assertEqual(cell_images["part_name"], "xl/CellImages1.XML")
        self.assertEqual(cell_images["content_type"], "application/vnd.ms-excel.cellimages+xml")
        self.assertEqual(cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages")
        self.assertEqual(cell_images["embed_rids_count"], 1)

    def test_triage_does_not_misattribute_workbook_rel_type_on_basename_collision(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(
                display_name="book.xlsx",
                data=_make_xlsx_with_cell_images_duplicate_basename_workbook_targets_root(),
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
        self.assertEqual(
            cell_images["workbook_rel_type"], "http://example.com/relationships/cellImages-root"
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
