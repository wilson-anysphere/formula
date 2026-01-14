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


def _make_xlsx_with_custom_functions() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>SUM(1,2)</f></c>
      <c r="A2"><f>CORP.ADDIN.FOO(1)</f></c>
    </row>
  </sheetData>
</worksheet>
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

    def test_private_mode_redacts_non_http_uri_schemes_in_diff_paths(self) -> None:
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
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="file:///C:/corp/secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Type="urn:corp:reltype"]@Type',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="//corp.example.com/share/secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="corp.example.com/share/secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="/Users/alice/secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="C:/Users/alice/secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="C:\\\\Users\\\\alice\\\\secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml.rels",
                                    "path": '/Relationships/Relationship[@Target="\\\\\\\\corp.example.com\\\\share\\\\secret.xlsx"]@Target',
                                    "kind": "attribute_changed",
                                },
                            ]
                        },
                    }
                },
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
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        top = report["steps"]["diff"]["details"]["top_differences"]
        self.assertEqual(len(top), 8)
        self.assertIn("sha256=", top[0]["path"])
        self.assertNotIn("file:///C:/corp/secret.xlsx", top[0]["path"])
        self.assertIn("sha256=", top[1]["path"])
        self.assertNotIn("urn:corp:reltype", top[1]["path"])
        self.assertIn("sha256=", top[2]["path"])
        self.assertNotIn("corp.example.com", top[2]["path"])
        self.assertIn("sha256=", top[3]["path"])
        self.assertNotIn("corp.example.com", top[3]["path"])
        self.assertIn("sha256=", top[4]["path"])
        self.assertNotIn("/Users/alice/secret.xlsx", top[4]["path"])
        self.assertIn("sha256=", top[5]["path"])
        self.assertNotIn("C:/Users/alice/secret.xlsx", top[5]["path"])
        self.assertIn("sha256=", top[6]["path"])
        self.assertNotIn("C:\\\\Users\\\\alice\\\\secret.xlsx", top[6]["path"])
        self.assertIn("sha256=", top[7]["path"])
        self.assertNotIn("\\\\\\\\corp.example.com\\\\share\\\\secret.xlsx", top[7]["path"])

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

    def test_private_mode_hashes_non_standard_function_names(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            data = _make_xlsx_with_custom_functions()
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

        functions = report.get("functions") or {}
        self.assertIsInstance(functions, dict)
        # Built-in Excel functions should remain readable.
        self.assertEqual(functions.get("SUM"), 1)

        custom = "CORP.ADDIN.FOO"
        expected_hash = hashlib.sha256(custom.encode("utf-8")).hexdigest()
        self.assertEqual(functions.get(f"sha256={expected_hash}"), 1)
        self.assertFalse(any(custom in k for k in functions.keys()))

    def test_private_mode_hashes_ignore_path_settings_in_diff_details(self) -> None:
        import tools.corpus.triage as triage_mod

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore_paths": [
                                "xr:uid",
                                "http://corp.example.com/ns",
                            ],
                            "top_differences": [],
                        },
                    }
                },
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
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        ignore_paths = report["steps"]["diff"]["details"]["ignore_paths"]
        self.assertIsInstance(ignore_paths, list)
        self.assertTrue(all(isinstance(p, str) and p.startswith("sha256=") for p in ignore_paths))
        self.assertFalse(any("corp.example.com" in p for p in ignore_paths))


if __name__ == "__main__":
    unittest.main()
