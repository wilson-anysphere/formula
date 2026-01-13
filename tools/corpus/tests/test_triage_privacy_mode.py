from __future__ import annotations

import hashlib
import io
import os
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


def _make_xlsx_with_custom_content_type() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.corp.example.cellimages+xml"/>
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

    def test_triage_workbook_private_mode_preserves_xlsb_extension(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            # A minimal ZIP is enough for the feature scan path; Rust triage is mocked.
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED):
                pass
            data = buf.getvalue()

            wb = WorkbookInput(display_name="book.xlsb", data=data)
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
        self.assertEqual(report["display_name"], f"workbook-{expected_sha[:16]}.xlsb")

    def test_triage_workbook_private_mode_redacts_custom_namespaces_in_diff_paths(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "top_differences": [
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml",
                                    "path": "/root@{http://corp.example.com/ns}attr",
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml",
                                    "path": "/root@{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                                    "kind": "attribute_changed",
                                },
                            ]
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            data = _make_xlsx_with_custom_relationship_uris()
            wb = WorkbookInput(display_name="book.xlsx", data=data)
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

        top = report["steps"]["diff"]["details"]["top_differences"]
        self.assertEqual(len(top), 2)
        self.assertIn("{sha256=", top[0]["path"])
        self.assertNotIn("corp.example.com", top[0]["path"])
        # Allowlisted OpenXML schema URIs should remain intact.
        self.assertIn("schemas.openxmlformats.org", top[1]["path"])

    def test_private_mode_hashes_non_github_run_url(self) -> None:
        import tools.corpus.triage as triage_mod

        original_env = os.environ.copy()
        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            os.environ["GITHUB_SERVER_URL"] = "https://github.corp.example.com"
            os.environ["GITHUB_REPOSITORY"] = "corp/repo"
            os.environ["GITHUB_RUN_ID"] = "123"

            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            data = _make_xlsx_with_custom_relationship_uris()
            report = triage_workbook(
                WorkbookInput(display_name="book.xlsx", data=data),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
                privacy_mode="private",
            )
        finally:
            os.environ.clear()
            os.environ.update(original_env)
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertIsInstance(report.get("run_url"), str)
        self.assertTrue(report["run_url"].startswith("sha256="))
        self.assertNotIn("github.corp.example.com", report["run_url"])

    def test_private_mode_hashes_custom_content_types(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            data = _make_xlsx_with_custom_content_type()
            report = triage_workbook(
                WorkbookInput(display_name="book.xlsx", data=data),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
                privacy_mode="private",
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        ct = report.get("cell_images", {}).get("content_type")
        self.assertIsInstance(ct, str)
        self.assertTrue(ct.startswith("sha256="))
        self.assertNotIn("corp.example", ct)


if __name__ == "__main__":
    unittest.main()
