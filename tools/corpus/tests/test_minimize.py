from __future__ import annotations

import io
import hashlib
import json
import sys
import tempfile
import unittest
from unittest import mock
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET

import tools.corpus.minimize as minimize_mod
from tools.corpus.util import WorkbookInput


class CorpusMinimizeTests(unittest.TestCase):
    def test_relationship_id_from_diff_path(self) -> None:
        self.assertEqual(
            minimize_mod._relationship_id_from_diff_path(  # noqa: SLF001 (unit test)
                '/Relationships/Relationship[@Id="rId3"]@Target'
            ),
            "rId3",
        )
        self.assertIsNone(minimize_mod._relationship_id_from_diff_path("/Relationships/Relationship"))  # noqa: SLF001

    def test_summarize_differences_counts_and_rel_ids(self) -> None:
        diffs = [
            {
                "severity": "CRITICAL",
                "part": "xl/workbook.xml",
                "path": "/workbook@{ns}attr",
                "kind": "attribute_changed",
            },
            {
                "severity": "CRITICAL",
                "part": "xl/_rels/workbook.xml.rels",
                "path": '/Relationships/Relationship[@Id="rId1"]@Target',
                "kind": "attribute_changed",
            },
            {
                "severity": "WARN",
                "part": "xl/theme/theme1.xml",
                "path": "/theme",
                "kind": "child_added",
            },
        ]

        per_part, critical_parts, rel_ids = minimize_mod.summarize_differences(diffs)

        self.assertEqual(critical_parts, ["xl/_rels/workbook.xml.rels", "xl/workbook.xml"])
        self.assertEqual(per_part["xl/workbook.xml"]["critical"], 1)
        self.assertEqual(per_part["xl/theme/theme1.xml"]["warning"], 1)
        self.assertEqual(rel_ids, {"xl/_rels/workbook.xml.rels": ["rId1"]})

    def test_minimize_workbook_reruns_when_truncated(self) -> None:
        calls: list[int] = []
        names: list[str] = []

        def fake_run_rust_triage(  # type: ignore[no-untyped-def]
            *_args, workbook_name: str, diff_limit: int, **_kwargs
        ):
            calls.append(diff_limit)
            names.append(workbook_name)
            full = [
                {
                    "severity": "CRITICAL",
                    "part": "xl/workbook.xml",
                    "path": "/workbook@{ns}attr",
                    "kind": "attribute_changed",
                },
                {
                    "severity": "CRITICAL",
                    "part": "xl/_rels/workbook.xml.rels",
                    "path": '/Relationships/Relationship[@Id="rId9"]@Target',
                    "kind": "attribute_changed",
                },
                {
                    "severity": "WARN",
                    "part": "xl/theme/theme1.xml",
                    "path": "/theme",
                    "kind": "child_added",
                },
            ]
            emitted = full[:1] if diff_limit < len(full) else full
            return {
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 2, "warning": 1, "info": 0, "total": 3},
                            "equal": False,
                            "top_differences": emitted,
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": False},
            }

        original = minimize_mod.triage_mod._run_rust_triage  # noqa: SLF001 (test patch)
        try:
            minimize_mod.triage_mod._run_rust_triage = fake_run_rust_triage  # type: ignore[assignment]
            summary = minimize_mod.minimize_workbook(
                WorkbookInput(display_name="book.xlsx", data=b"dummy"),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=1,
            )
        finally:
            minimize_mod.triage_mod._run_rust_triage = original  # type: ignore[assignment]

        # First run truncated (diff_limit=1) then rerun with diff_limit=total (=3).
        self.assertEqual(calls, [1, 3])
        self.assertEqual(names, ["book.xlsx", "book.xlsx"])
        self.assertEqual(summary["critical_parts"], ["xl/_rels/workbook.xml.rels", "xl/workbook.xml"])
        self.assertEqual(summary["rels_critical_ids"], {"xl/_rels/workbook.xml.rels": ["rId9"]})

    def test_minimize_workbook_does_not_rerun_with_parts_with_diffs(self) -> None:
        calls: list[int] = []

        def fake_run_rust_triage(  # type: ignore[no-untyped-def]
            *_args, workbook_name: str, diff_limit: int, **_kwargs
        ):
            calls.append(diff_limit)
            self.assertEqual(workbook_name, "book.xlsx")
            full = [
                {
                    "severity": "CRITICAL",
                    "part": "xl/workbook.xml",
                    "path": "/workbook@{ns}attr",
                    "kind": "attribute_changed",
                },
                {
                    "severity": "WARN",
                    "part": "xl/theme/theme1.xml",
                    "path": "/theme",
                    "kind": "child_added",
                },
                {
                    "severity": "WARN",
                    "part": "xl/theme/theme1.xml",
                    "path": "/theme",
                    "kind": "child_added",
                },
            ]
            emitted = full[:diff_limit]
            return {
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 1, "warning": 2, "info": 0, "total": 3},
                            "equal": False,
                            # Newer triage helpers include per-part summaries so callers don't
                            # need to request every diff entry.
                            "critical_parts": ["xl/workbook.xml"],
                            "parts_with_diffs": [
                                {
                                    "part": "xl/workbook.xml",
                                    "group": "other",
                                    "critical": 1,
                                    "warning": 0,
                                    "info": 0,
                                    "total": 1,
                                },
                                {
                                    "part": "xl/theme/theme1.xml",
                                    "group": "other",
                                    "critical": 0,
                                    "warning": 2,
                                    "info": 0,
                                    "total": 2,
                                },
                            ],
                            "top_differences": emitted,
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": False},
            }

        original = minimize_mod.triage_mod._run_rust_triage  # noqa: SLF001 (test patch)
        try:
            minimize_mod.triage_mod._run_rust_triage = fake_run_rust_triage  # type: ignore[assignment]
            summary = minimize_mod.minimize_workbook(
                WorkbookInput(display_name="book.xlsx", data=b"dummy"),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=1,
            )
        finally:
            minimize_mod.triage_mod._run_rust_triage = original  # type: ignore[assignment]

        # No rerun with `diff_limit=total` should occur.
        self.assertEqual(calls, [1])
        self.assertEqual(summary["critical_parts"], ["xl/workbook.xml"])
        self.assertEqual(summary["part_counts"]["xl/theme/theme1.xml"]["warning"], 2)
        self.assertTrue(
            any(p["part"] == "xl/theme/theme1.xml" for p in summary["parts_with_diffs"])
        )

    def test_minimize_workbook_reruns_only_for_all_critical_rels_diffs(self) -> None:
        calls: list[int] = []

        def fake_run_rust_triage(  # type: ignore[no-untyped-def]
            *_args, workbook_name: str, diff_limit: int, **_kwargs
        ):
            calls.append(diff_limit)
            self.assertEqual(workbook_name, "book.xlsx")
            full = [
                {
                    "severity": "CRITICAL",
                    "part": "xl/_rels/workbook.xml.rels",
                    "path": '/Relationships/Relationship[@Id="rId1"]@Target',
                    "kind": "attribute_changed",
                },
                {
                    "severity": "CRITICAL",
                    "part": "xl/_rels/workbook.xml.rels",
                    "path": '/Relationships/Relationship[@Id="rId9"]@Target',
                    "kind": "attribute_changed",
                },
                {
                    "severity": "WARN",
                    "part": "xl/theme/theme1.xml",
                    "path": "/theme",
                    "kind": "child_added",
                },
            ]
            emitted = full[:diff_limit]
            return {
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 2, "warning": 1, "info": 0, "total": 3},
                            "equal": False,
                            "critical_parts": ["xl/_rels/workbook.xml.rels"],
                            "parts_with_diffs": [
                                {
                                    "part": "xl/_rels/workbook.xml.rels",
                                    "group": "rels",
                                    "critical": 2,
                                    "warning": 0,
                                    "info": 0,
                                    "total": 2,
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
                            "top_differences": emitted,
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": False},
            }

        original = minimize_mod.triage_mod._run_rust_triage  # noqa: SLF001 (test patch)
        try:
            minimize_mod.triage_mod._run_rust_triage = fake_run_rust_triage  # type: ignore[assignment]
            summary = minimize_mod.minimize_workbook(
                WorkbookInput(display_name="book.xlsx", data=b"dummy"),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=1,
            )
        finally:
            minimize_mod.triage_mod._run_rust_triage = original  # type: ignore[assignment]

        # First run truncated (diff_limit=1). We only need CRITICAL diffs to extract rIds, so
        # rerun with diff_limit=critical (=2), not total (=3).
        self.assertEqual(calls, [1, 2])
        self.assertEqual(
            summary["rels_critical_ids"], {"xl/_rels/workbook.xml.rels": ["rId1", "rId9"]}
        )

    def test_minimize_workbook_includes_critical_part_hashes(self) -> None:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/workbook.xml", b"abc")
        wb = WorkbookInput(display_name="book.xlsx", data=buf.getvalue())

        def fake_run_rust_triage(*_args, **_kwargs):  # type: ignore[no-untyped-def]
            return {
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 1, "warning": 0, "info": 0, "total": 1},
                            "equal": False,
                            "top_differences": [
                                {
                                    "severity": "CRITICAL",
                                    "part": "xl/workbook.xml",
                                    "path": "/workbook",
                                    "kind": "binary_diff",
                                }
                            ],
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": False},
            }

        original = minimize_mod.triage_mod._run_rust_triage  # noqa: SLF001
        try:
            minimize_mod.triage_mod._run_rust_triage = fake_run_rust_triage  # type: ignore[assignment]
            summary = minimize_mod.minimize_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=10,
            )
        finally:
            minimize_mod.triage_mod._run_rust_triage = original  # type: ignore[assignment]

        expected_hash = hashlib.sha256(b"abc").hexdigest()
        self.assertEqual(
            summary["critical_part_hashes"]["xl/workbook.xml"]["sha256"],
            expected_hash,
        )

    def test_main_writes_summary_json(self) -> None:
        def fake_build_rust_helper() -> Path:  # type: ignore[no-untyped-def]
            return Path("noop")

        def fake_run_rust_triage(*_args, workbook_name: str, **_kwargs):  # type: ignore[no-untyped-def]
            self.assertEqual(workbook_name, "book.xlsx")
            return {
                "steps": {
                    "diff": {
                        "status": "ok",
                        "details": {
                            "ignore": [],
                            "counts": {"critical": 0, "warning": 0, "info": 0, "total": 0},
                            "equal": True,
                            "top_differences": [],
                        },
                    }
                },
                "result": {"open_ok": True, "round_trip_ok": True},
            }

        orig_build = minimize_mod.triage_mod._build_rust_helper  # noqa: SLF001
        orig_run = minimize_mod.triage_mod._run_rust_triage  # noqa: SLF001
        try:
            minimize_mod.triage_mod._build_rust_helper = fake_build_rust_helper  # type: ignore[assignment]
            minimize_mod.triage_mod._run_rust_triage = fake_run_rust_triage  # type: ignore[assignment]

            with tempfile.TemporaryDirectory() as tmpdir:
                in_path = Path(tmpdir) / "book.xlsx"
                in_path.write_bytes(b"dummy")
                out_path = Path(tmpdir) / "summary.json"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.minimize",
                        "--input",
                        str(in_path),
                        "--out",
                        str(out_path),
                    ]
                    # Avoid polluting unit test output; the CLI is intentionally chatty.
                    buf = io.StringIO()
                    with mock.patch("sys.stdout", buf):
                        rc = minimize_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                data = json.loads(out_path.read_text(encoding="utf-8"))
                self.assertEqual(data["display_name"], "book.xlsx")
                self.assertEqual(data["diff_counts"]["critical"], 0)
        finally:
            minimize_mod.triage_mod._build_rust_helper = orig_build  # type: ignore[assignment]
            minimize_mod.triage_mod._run_rust_triage = orig_run  # type: ignore[assignment]

    def test_prune_xlsx_parts_removes_content_types_and_rels_references(self) -> None:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr(
                "[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
""",
            )
            z.writestr(
                "_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""",
            )
            z.writestr(
                "xl/workbook.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
""",
            )
            z.writestr(
                "xl/_rels/workbook.xml.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
""",
            )
            z.writestr(
                "xl/styles.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
""",
            )
            z.writestr(
                "xl/worksheets/sheet1.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>
""",
            )
            z.writestr(
                "xl/theme/theme1.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
""",
            )

        original_bytes = buf.getvalue()
        keep = {
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/worksheets/sheet1.xml",
        }
        pruned = minimize_mod.prune_xlsx_parts(original_bytes, keep_parts=keep)

        with zipfile.ZipFile(io.BytesIO(pruned), "r") as z:
            names = {info.filename for info in z.infolist() if not info.is_dir()}
            self.assertNotIn("xl/theme/theme1.xml", names)

            ct = ET.fromstring(z.read("[Content_Types].xml"))
            overrides = [
                el.attrib.get("PartName", "")
                for el in ct
                if el.tag.split("}")[-1] == "Override"
            ]
            self.assertNotIn("/xl/theme/theme1.xml", overrides)

            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            targets = [
                el.attrib.get("Target", "")
                for el in rels
                if el.tag.split("}")[-1] == "Relationship"
            ]
            self.assertNotIn("theme/theme1.xml", targets)

    def test_required_core_parts_detects_xlsb_workbook_bin(self) -> None:
        parts = {
            "_rels/.rels": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="xl/workbook.bin"/>
</Relationships>""",
            "xl/workbook.bin": b"BIN",
            "xl/_rels/workbook.bin.rels": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.bin"/>
</Relationships>""",
            "xl/worksheets/sheet1.bin": b"SHEET",
        }

        required = minimize_mod._required_core_parts(parts)  # noqa: SLF001 (unit test)
        self.assertIn("_rels/.rels", required)
        self.assertIn("xl/workbook.bin", required)
        self.assertIn("xl/_rels/workbook.bin.rels", required)
        self.assertIn("xl/worksheets/sheet1.bin", required)

    def test_minimize_workbook_package_greedy_removal_stops_on_changed_critical_set(self) -> None:
        # Build a tiny (synthetic) XLSX-like zip with a removable theme part + a large binary part.
        # We'll mock `minimize_workbook` so we can unit test the greedy removal logic without
        # invoking Rust.
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr(
                "[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/big.bin" ContentType="application/octet-stream"/>
</Types>
""",
            )
            z.writestr(
                "_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""",
            )
            z.writestr(
                "xl/workbook.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
""",
            )
            z.writestr(
                "xl/_rels/workbook.xml.rels",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
""",
            )
            z.writestr(
                "xl/styles.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
""",
            )
            z.writestr(
                "xl/worksheets/sheet1.xml",
                """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>
""",
            )
            z.writestr("xl/theme/theme1.xml", b"x" * 2000)
            z.writestr("xl/theme/_rels/theme1.xml.rels", b"y" * 50)
            z.writestr("xl/big.bin", b"z" * 1500)

        data = buf.getvalue()
        wb = WorkbookInput(display_name="book.xlsx", data=data)

        baseline = {
            "open_ok": True,
            "round_trip_ok": False,
            "critical_parts": ["xl/workbook.xml"],
            "diff_counts": {"critical": 1, "warning": 0, "info": 0, "total": 1},
        }

        def fake_minimize_workbook(workbook: WorkbookInput, **_kwargs):  # type: ignore[no-untyped-def]
            with zipfile.ZipFile(io.BytesIO(workbook.data), "r") as z:
                names = {info.filename for info in z.infolist() if not info.is_dir()}
            # Pretend removing `xl/big.bin` changes the critical diff signature, so the minimizer
            # should reject that removal.
            if "xl/big.bin" not in names:
                return {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "critical_parts": ["xl/workbook.xml", "xl/_rels/workbook.xml.rels"],
                    "diff_counts": {"critical": 2, "warning": 0, "info": 0, "total": 2},
                }
            return baseline

        original = minimize_mod.minimize_workbook  # noqa: SLF001
        try:
            minimize_mod.minimize_workbook = fake_minimize_workbook  # type: ignore[assignment]
            with tempfile.TemporaryDirectory() as tmpdir:
                out_xlsx = Path(tmpdir) / "min.xlsx"
                _summary, removed, out_bytes = minimize_mod.minimize_workbook_package(
                    wb,
                    rust_exe=Path("noop"),
                    diff_ignore=set(),
                    diff_limit=0,
                    out_xlsx=out_xlsx,
                    max_steps=50,
                    baseline=baseline,
                )
        finally:
            minimize_mod.minimize_workbook = original  # type: ignore[assignment]

        self.assertIn("xl/theme/theme1.xml", removed)
        self.assertIn("xl/theme/_rels/theme1.xml.rels", removed)
        self.assertNotIn("xl/big.bin", removed)

        with zipfile.ZipFile(io.BytesIO(out_bytes), "r") as z:
            names = {info.filename for info in z.infolist() if not info.is_dir()}
            self.assertNotIn("xl/theme/theme1.xml", names)
            self.assertNotIn("xl/theme/_rels/theme1.xml.rels", names)
            self.assertIn("xl/big.bin", names)

            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            targets = [
                el.attrib.get("Target", "")
                for el in rels
                if el.tag.split("}")[-1] == "Relationship"
            ]
            self.assertNotIn("theme/theme1.xml", targets)


if __name__ == "__main__":
    unittest.main()
