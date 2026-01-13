from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class ComparePartialDatasetsTests(unittest.TestCase):
    def test_missing_expected_does_not_crash_and_emits_report(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {"id": "case-a", "formula": "=1+1", "outputCell": "C1", "inputs": []},
                    {
                        "id": "case-b",
                        "formula": "=2+2",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["tag-b"],
                        "description": "Case B description",
                    },
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {
                    "kind": "excel",
                    "note": "synthetic test fixture",
                    "syntheticSource": {
                        "kind": "formula-engine",
                        "version": "unit-test",
                        "os": "linux",
                        "arch": "x86_64",
                        "caseSet": "unit-test",
                    },
                    "patches": [{"version": "16.0", "build": "unit-test", "operatingSystem": "windows"}],
                },
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 2}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-a", "result": {"t": "n", "v": 2}},
                    {
                        "caseId": "case-b",
                        "result": {"t": "n", "v": 4},
                        "address": "C1",
                        "displayText": "4",
                    },
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "1.0",
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}")

            self.assertTrue(report_path.is_file(), "compare.py did not write a report JSON")
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(report.get("expectedSource"), expected_payload["source"])
            self.assertEqual(report.get("actualSource"), actual_payload["source"])
            self.assertEqual(report["summary"]["reasonCounts"]["missing-expected"], 1)
            self.assertEqual(report["summary"]["casesPath"], str(cases_path))
            self.assertEqual(report["summary"]["expectedPath"], str(expected_path))
            self.assertEqual(report["summary"]["actualPath"], str(actual_path))
            self.assertEqual(report["summary"]["expectedDatasetKind"], "synthetic")
            self.assertEqual(report["summary"]["expectedDatasetHasPatches"], True)
            self.assertEqual(report["summary"]["expectedDatasetPatchEntryCount"], 1)
            mismatches = report.get("mismatches", [])
            self.assertIsInstance(mismatches, list)
            self.assertEqual(len(mismatches), 1)
            self.assertEqual(mismatches[0]["caseId"], "case-b")
            self.assertEqual(mismatches[0]["tags"], ["tag-b"])
            self.assertEqual(mismatches[0]["actual"], {"t": "n", "v": 4})
            self.assertEqual(mismatches[0]["outputCell"], "C1")
            self.assertEqual(mismatches[0]["description"], "Case B description")
            self.assertEqual(mismatches[0]["actualAddress"], "C1")
            self.assertEqual(mismatches[0]["actualDisplayText"], "4")


class CompareTagToleranceTests(unittest.TestCase):
    def test_tag_specific_tolerance_allows_small_numeric_diffs(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=1/3",
                        "inputs": [],
                        "tags": ["loose"],
                    }
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 0.3333333333333333}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 0.333333}}],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            # Without tag-specific tolerances this should fail (default abs/rel tolerances are 1e-9).
            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "0.0",
                ],
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(
                proc.returncode,
                0,
                f"Expected compare.py to fail without tag tolerance.\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "0.0",
                    "--tag-abs-tol",
                    "loose=1e-6",
                    "--tag-rel-tol",
                    "loose=1e-6",
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode} with tag tolerances\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(report["summary"]["mismatches"], 0)
            self.assertEqual(report["summary"]["tagAbsTol"], {"loose": 1e-6})
            self.assertEqual(report["summary"]["tagRelTol"], {"loose": 1e-6})
            self.assertEqual(report["summary"]["casesPath"], str(cases_path))
            self.assertEqual(report["summary"]["expectedPath"], str(expected_path))
            self.assertEqual(report["summary"]["actualPath"], str(actual_path))


class CompareMismatchDetailTests(unittest.TestCase):
    def test_number_mismatch_includes_tolerances_and_diffs(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            expected_value = 0.3333333333333333
            actual_value = 0.333333

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=1/3",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["num"],
                    }
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "n", "v": expected_value},
                        "address": "C1",
                        "displayText": str(expected_value),
                    }
                ],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "n", "v": actual_value},
                        "address": "C1",
                        "displayText": str(actual_value),
                    }
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    # Allow mismatches so we can inspect the report.
                    "--max-mismatch-rate",
                    "1.0",
                    "--abs-tol",
                    "1e-9",
                    "--rel-tol",
                    "1e-9",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            mismatches = report.get("mismatches", [])
            self.assertIsInstance(mismatches, list)
            self.assertEqual(len(mismatches), 1)

            m = mismatches[0]
            self.assertEqual(m.get("reason"), "number-mismatch")
            self.assertEqual(m.get("tags"), ["num"])
            self.assertEqual(m.get("outputCell"), "C1")
            self.assertEqual(m.get("absTol"), 1e-9)
            self.assertEqual(m.get("relTol"), 1e-9)
            self.assertEqual(m.get("expectedAddress"), "C1")
            self.assertEqual(m.get("actualAddress"), "C1")
            self.assertEqual(m.get("expectedDisplayText"), str(expected_value))
            self.assertEqual(m.get("actualDisplayText"), str(actual_value))

            abs_diff = abs(expected_value - actual_value)
            denom = max(abs(expected_value), abs(actual_value))
            rel_diff = abs_diff / denom if denom else None

            self.assertAlmostEqual(m.get("absDiff"), abs_diff, places=15)
            self.assertAlmostEqual(m.get("relDiff"), rel_diff, places=15)


class CompareArrayMismatchDetailTests(unittest.TestCase):
    def test_array_mismatch_includes_cell_position(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=SEQUENCE(2,2)",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["arr"],
                    }
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "address": "C1:D2",
                        "displayText": "1",
                        "result": {
                            "t": "arr",
                            "rows": [
                                [
                                    {"t": "n", "v": 1},
                                    {"t": "n", "v": 2},
                                ],
                                [
                                    {"t": "n", "v": 3},
                                    {"t": "n", "v": 4},
                                ],
                            ],
                        },
                    }
                ],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "address": "C1:D2",
                        "displayText": "1",
                        "result": {
                            "t": "arr",
                            "rows": [
                                [
                                    {"t": "n", "v": 1},
                                    {"t": "n", "v": 2},
                                ],
                                [
                                    {"t": "n", "v": 999},
                                    {"t": "n", "v": 4},
                                ],
                            ],
                        },
                    }
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "1.0",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            mismatches = report.get("mismatches", [])
            self.assertIsInstance(mismatches, list)
            self.assertEqual(len(mismatches), 1)
            m = mismatches[0]
            self.assertEqual(m.get("reason"), "array-mismatch:number-mismatch")
            self.assertEqual(m.get("tags"), ["arr"])
            detail = m.get("mismatchDetail")
            self.assertEqual(
                detail,
                {
                    "row": 1,
                    "col": 0,
                    "reason": "number-mismatch",
                    "detail": None,
                    "expected": {"t": "n", "v": 3},
                    "actual": {"t": "n", "v": 999},
                },
            )


class CompareArrayShapeMismatchDetailTests(unittest.TestCase):
    def test_array_row_count_mismatch_includes_shape(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=SEQUENCE(2,2)",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["arr"],
                    }
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "address": "C1:D2",
                        "displayText": "1",
                        "result": {
                            "t": "arr",
                            "rows": [
                                [
                                    {"t": "n", "v": 1},
                                    {"t": "n", "v": 2},
                                ],
                                [
                                    {"t": "n", "v": 3},
                                    {"t": "n", "v": 4},
                                ],
                            ],
                        },
                    }
                ],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "address": "C1:D1",
                        "displayText": "1",
                        "result": {
                            "t": "arr",
                            "rows": [
                                [
                                    {"t": "n", "v": 1},
                                    {"t": "n", "v": 2},
                                ]
                            ],
                        },
                    }
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "1.0",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            mismatches = report.get("mismatches", [])
            self.assertIsInstance(mismatches, list)
            self.assertEqual(len(mismatches), 1)
            m = mismatches[0]
            self.assertEqual(m.get("reason"), "array-shape-mismatch")
            self.assertEqual(m.get("tags"), ["arr"])
            self.assertEqual(m.get("mismatchDetail"), {"expectedRows": 2, "actualRows": 1})


class CompareDuplicateCaseIdTests(unittest.TestCase):
    def test_duplicate_case_ids_fail_fast(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {"id": "case-a", "formula": "=1", "inputs": []},
                    {"id": "case-b", "formula": "=2", "inputs": []},
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-a", "result": {"t": "n", "v": 1}},
                    {"caseId": "case-b", "result": {"t": "n", "v": 2}},
                ],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            # Duplicate case-a entry.
            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-a", "result": {"t": "n", "v": 1}},
                    {"caseId": "case-a", "result": {"t": "n", "v": 1}},
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--max-mismatch-rate",
                    "1.0",
                ],
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(proc.returncode, 0)
            self.assertIn("duplicate caseId", proc.stderr)


class CompareDryRunTests(unittest.TestCase):
    def test_dry_run_does_not_write_report(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {"id": "case-a", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": ["t1"]},
                    {"id": "case-b", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": ["t2"]},
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-a", "result": {"t": "n", "v": 2}},
                    {"caseId": "case-b", "result": {"t": "n", "v": 4}},
                ],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "note": "synthetic test fixture"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-a", "result": {"t": "n", "v": 2}},
                    {"caseId": "case-b", "result": {"t": "n", "v": 4}},
                ],
            }
            actual_path.write_text(
                json.dumps(actual_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(compare_py),
                    "--cases",
                    str(cases_path),
                    "--expected",
                    str(expected_path),
                    "--actual",
                    str(actual_path),
                    "--report",
                    str(report_path),
                    "--include-tag",
                    "t1",
                    "--dry-run",
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compare.py dry-run exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            self.assertFalse(report_path.exists(), "compare.py dry-run should not write report.json")
            self.assertIn("Dry run: compare.py", proc.stdout)
            self.assertIn("cases selected: 1", proc.stdout)


if __name__ == "__main__":
    unittest.main()
