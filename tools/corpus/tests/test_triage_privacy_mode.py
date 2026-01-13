from __future__ import annotations

import hashlib
import io
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import triage_workbook
from tools.corpus.util import WorkbookInput, sha256_hex


def _make_xlsx_with_custom_relationship_uris() -> bytes:
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
</ci:cellImages>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId99"
                Type="http://corp.example.com/relationships/cellImages"
                Target="cellimages.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://corp.example.com/relationships/image" Target="media/image2.png"/>
</Relationships>
""",
        )
    return buf.getvalue()


class TriagePrivacyModeTests(unittest.TestCase):
    def test_triage_workbook_private_mode_anonymizes_display_name_and_hashes_custom_uris(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            data = _make_xlsx_with_custom_relationship_uris()
            wb = WorkbookInput(display_name="sensitive-filename.xlsx", data=data)
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

        expected_sha = sha256_hex(data)
        self.assertEqual(report["display_name"], f"workbook-{expected_sha[:16]}.xlsx")

        cell_images = report["cell_images"]

        # Allowlisted schema URIs should be preserved.
        self.assertEqual(
            cell_images["root_namespace"],
            "http://schemas.microsoft.com/office/spreadsheetml/2024/cellimages",
        )

        # Custom relationship type URIs should be hashed.
        expected_rel_hash = hashlib.sha256(
            "http://corp.example.com/relationships/cellImages".encode("utf-8")
        ).hexdigest()
        self.assertEqual(cell_images["workbook_rel_type"], f"sha256={expected_rel_hash}")

        rels_types = cell_images["rels_types"]
        self.assertIn(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            rels_types,
        )
        self.assertTrue(any(v.startswith("sha256=") for v in rels_types))
        self.assertFalse(any("corp.example.com" in v for v in rels_types))


if __name__ == "__main__":
    unittest.main()

