from __future__ import annotations

import hashlib
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


class ComparePrivacyModeTests(unittest.TestCase):
    def test_privacy_mode_hashes_paths_in_report_summary(self) -> None:
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
                "cases": [{"id": "case-a", "formula": "=1+1", "outputCell": "C1", "inputs": []}],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            secret_patch_path = str(tmp_path / "patch-output.json")
            relative_patch_path = "tools/excel-oracle/odd_coupon_boundary_cases.json"
            file_patch_path = "file:///home/alice/patch-output.json"
            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {
                    "kind": "excel",
                    "note": "synthetic test fixture",
                    "patches": [
                        {
                            "version": "16.0",
                            "build": "unit-test",
                            "operatingSystem": "windows",
                            "caseSet": {"path": secret_patch_path},
                        },
                        {
                            "version": "16.0",
                            "build": "unit-test",
                            "operatingSystem": "windows",
                            "caseSet": {"path": relative_patch_path},
                        },
                        {
                            "version": "16.0",
                            "build": "unit-test",
                            "operatingSystem": "windows",
                            "caseSet": {"path": file_patch_path},
                        },
                    ],
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
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 2}}],
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
                    "--privacy-mode",
                    "private",
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(
                report["summary"]["casesPath"],
                f"sha256={hashlib.sha256(str(cases_path).encode('utf-8')).hexdigest()}",
            )
            self.assertEqual(
                report["summary"]["expectedPath"],
                f"sha256={hashlib.sha256(str(expected_path).encode('utf-8')).hexdigest()}",
            )
            self.assertEqual(
                report["summary"]["actualPath"],
                f"sha256={hashlib.sha256(str(actual_path).encode('utf-8')).hexdigest()}",
            )

            # Defense in depth: redact embedded paths inside expected/actual source metadata too.
            redacted_patch_path = report["expectedSource"]["patches"][0]["caseSet"]["path"]
            self.assertEqual(
                redacted_patch_path,
                f"sha256={hashlib.sha256(secret_patch_path.encode('utf-8')).hexdigest()}",
            )
            self.assertEqual(
                report["expectedSource"]["patches"][1]["caseSet"]["path"],
                relative_patch_path,
            )
            self.assertEqual(
                report["expectedSource"]["patches"][2]["caseSet"]["path"],
                f"sha256={hashlib.sha256(file_patch_path.encode('utf-8')).hexdigest()}",
            )

    def test_privacy_mode_hashes_error_detail_strings(self) -> None:
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
                "cases": [{"id": "case-a", "formula": "=1+1", "outputCell": "C1", "inputs": []}],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 2}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            secret_detail = "Sensitive path: /home/alice/secret.xlsx"
            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "e", "v": "#VALUE!", "detail": secret_detail},
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
                    "--privacy-mode",
                    "private",
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
            actual_value = mismatches[0].get("actual")
            self.assertIsInstance(actual_value, dict)

            expected_hash = hashlib.sha256(secret_detail.encode("utf-8")).hexdigest()
            self.assertEqual(actual_value.get("detail"), f"sha256={expected_hash}")
            self.assertNotIn(secret_detail, json.dumps(report))

    def test_privacy_mode_redacts_paths_in_source_metadata_strings(self) -> None:
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
                "cases": [{"id": "case-a", "formula": "=1+1", "outputCell": "C1", "inputs": []}],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 2}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            secret_engine_path = r"C:\Users\Alice\engine.exe"
            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine", "engineBinary": secret_engine_path},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 3}}],
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
                    "--privacy-mode",
                    "private",
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
            actual_source = report.get("actualSource")
            self.assertIsInstance(actual_source, dict)
            expected_hash = hashlib.sha256(secret_engine_path.encode("utf-8")).hexdigest()
            self.assertEqual(actual_source.get("engineBinary"), f"sha256={expected_hash}")
            self.assertNotIn(secret_engine_path, json.dumps(report))

    def test_privacy_mode_hashes_domain_like_descriptions(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            secret_desc = "See corp.example.com for details"
            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=1+1",
                        "outputCell": "C1",
                        "inputs": [],
                        "description": secret_desc,
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
                "source": {"kind": "excel"},
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 3}}],
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
                    "--privacy-mode",
                    "private",
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

            expected_hash = hashlib.sha256(secret_desc.encode("utf-8")).hexdigest()
            self.assertEqual(mismatches[0].get("description"), f"sha256={expected_hash}")
            self.assertNotIn(secret_desc, json.dumps(report))

    def test_privacy_mode_hashes_namespace_like_tags_and_tag_tolerances(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            udf_tag = "CORP.ADDIN.FOO"
            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-sum",
                        "formula": "=SUM(1,2)",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["SUM"],
                    },
                    {
                        "id": "case-udf",
                        "formula": "=CORP.ADDIN.FOO(1)",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": [udf_tag],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "n", "v": 3}},
                    {"caseId": "case-udf", "result": {"t": "n", "v": 1}},
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "e", "v": "#NAME?"}},
                    {"caseId": "case-udf", "result": {"t": "e", "v": "#NAME?"}},
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
                    "--privacy-mode",
                    "private",
                    "--max-mismatch-rate",
                    "1.0",
                    "--tag-abs-tol",
                    f"{udf_tag}=1e-6",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(
                    f"compare.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            report = json.loads(report_path.read_text(encoding="utf-8"))
            mismatches = report.get("mismatches")
            self.assertIsInstance(mismatches, list)
            by_id = {m.get("caseId"): m for m in mismatches if isinstance(m, dict)}

            udf_hash = hashlib.sha256(udf_tag.encode("utf-8")).hexdigest()
            self.assertEqual(by_id["case-sum"]["tags"], ["SUM"])
            self.assertEqual(by_id["case-udf"]["tags"], [f"sha256={udf_hash}"])

            tag_summary = report.get("summary", {}).get("tagSummary")
            self.assertIsInstance(tag_summary, list)
            tags = {row.get("tag") for row in tag_summary if isinstance(row, dict)}
            self.assertIn("SUM", tags)
            self.assertNotIn(udf_tag, tags)
            self.assertIn(f"sha256={udf_hash}", tags)

            tag_abs_tol = report.get("summary", {}).get("tagAbsTol")
            self.assertIsInstance(tag_abs_tol, dict)
            self.assertNotIn(udf_tag, tag_abs_tol)
            self.assertEqual(tag_abs_tol.get(f"sha256={udf_hash}"), 1e-6)

    def test_privacy_mode_hashes_unknown_uppercase_tags(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            udf_tag = "MYUDF"
            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {"id": "case-sum", "formula": "=SUM(1,2)", "outputCell": "C1", "inputs": [], "tags": ["SUM"]},
                    {
                        "id": "case-udf",
                        "formula": f"={udf_tag}(1)",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": [udf_tag],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "n", "v": 3}},
                    {"caseId": "case-udf", "result": {"t": "n", "v": 1}},
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "e", "v": "#NAME?"}},
                    {"caseId": "case-udf", "result": {"t": "e", "v": "#NAME?"}},
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
                    "--privacy-mode",
                    "private",
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
            mismatches = report.get("mismatches")
            self.assertIsInstance(mismatches, list)
            by_id = {m.get("caseId"): m for m in mismatches if isinstance(m, dict)}

            udf_hash = hashlib.sha256(udf_tag.encode("utf-8")).hexdigest()
            self.assertEqual(by_id["case-sum"]["tags"], ["SUM"])
            self.assertEqual(by_id["case-udf"]["tags"], [f"sha256={udf_hash}"])

            tag_summary = report.get("summary", {}).get("tagSummary")
            self.assertIsInstance(tag_summary, list)
            tags = {row.get("tag") for row in tag_summary if isinstance(row, dict)}
            self.assertIn("SUM", tags)
            self.assertNotIn(udf_tag, tags)
            self.assertIn(f"sha256={udf_hash}", tags)

    def test_privacy_mode_hashes_path_like_string_values(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            secret_value = "file:///home/alice/secret.xlsx"
            other_value = "file:///home/alice/other.xlsx"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=A1",
                        "outputCell": "C1",
                        "inputs": [{"cell": "A1", "value": secret_value}],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "s", "v": secret_value},
                        "address": "C1",
                        "displayText": secret_value,
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "s", "v": other_value},
                        "address": "C1",
                        "displayText": other_value,
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
                    "--privacy-mode",
                    "private",
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

            secret_hash = hashlib.sha256(secret_value.encode("utf-8")).hexdigest()
            other_hash = hashlib.sha256(other_value.encode("utf-8")).hexdigest()

            self.assertEqual(m.get("expected"), {"t": "s", "v": f"sha256={secret_hash}"})
            self.assertEqual(m.get("actual"), {"t": "s", "v": f"sha256={other_hash}"})
            self.assertEqual(m.get("expectedDisplayText"), f"sha256={secret_hash}")
            self.assertEqual(m.get("actualDisplayText"), f"sha256={other_hash}")

            inputs = m.get("inputs")
            self.assertIsInstance(inputs, list)
            self.assertEqual(inputs[0].get("value"), f"sha256={secret_hash}")

            # Defense in depth: ensure the raw strings do not appear in the report at all.
            report_text = json.dumps(report)
            self.assertNotIn(secret_value, report_text)
            self.assertNotIn(other_value, report_text)

    def test_privacy_mode_hashes_domain_like_string_values(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            secret_value = "corp.example.com"
            other_value = "other.internal"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": "=A1",
                        "outputCell": "C1",
                        "inputs": [{"cell": "A1", "value": secret_value}],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "s", "v": secret_value},
                        "address": "C1",
                        "displayText": secret_value,
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [
                    {
                        "caseId": "case-a",
                        "result": {"t": "s", "v": other_value},
                        "address": "C1",
                        "displayText": other_value,
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
                    "--privacy-mode",
                    "private",
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

            secret_hash = hashlib.sha256(secret_value.encode("utf-8")).hexdigest()
            other_hash = hashlib.sha256(other_value.encode("utf-8")).hexdigest()

            self.assertEqual(m.get("expected"), {"t": "s", "v": f"sha256={secret_hash}"})
            self.assertEqual(m.get("actual"), {"t": "s", "v": f"sha256={other_hash}"})
            self.assertEqual(m.get("expectedDisplayText"), f"sha256={secret_hash}")
            self.assertEqual(m.get("actualDisplayText"), f"sha256={other_hash}")

            inputs = m.get("inputs")
            self.assertIsInstance(inputs, list)
            self.assertEqual(inputs[0].get("value"), f"sha256={secret_hash}")

            report_text = json.dumps(report)
            self.assertNotIn(secret_value, report_text)
            self.assertNotIn(other_value, report_text)

    def test_privacy_mode_hashes_path_like_string_literals_in_formulas(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            secret_value = "file:///home/alice/secret.xlsx"
            other_value = "file:///home/alice/other.xlsx"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": f'=\"{secret_value}\"',
                        "outputCell": "C1",
                        "inputs": [],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "s", "v": secret_value}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "s", "v": other_value}}],
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
                    "--privacy-mode",
                    "private",
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

            secret_hash = hashlib.sha256(secret_value.encode("utf-8")).hexdigest()
            redacted_formula = m.get("formula")
            self.assertEqual(redacted_formula, f'=\"sha256={secret_hash}\"')

            report_text = json.dumps(report)
            self.assertNotIn(secret_value, report_text)
            self.assertNotIn(other_value, report_text)

    def test_privacy_mode_hashes_path_like_bracketed_segments_in_formulas(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            secret_ref = "corp.example.com.xlsx"
            formula = f"=[{secret_ref}]Sheet1!A1"

            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {
                        "id": "case-a",
                        "formula": formula,
                        "outputCell": "C1",
                        "inputs": [],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 1}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": "case-a", "result": {"t": "n", "v": 2}}],
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
                    "--privacy-mode",
                    "private",
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

            ref_hash = hashlib.sha256(secret_ref.encode("utf-8")).hexdigest()
            self.assertEqual(m.get("formula"), f"=[sha256={ref_hash}]Sheet1!A1")
            self.assertNotIn(secret_ref, json.dumps(report))

    def test_privacy_mode_hashes_domain_like_case_ids(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            sensitive_case_id = "corp.example.com-case-a"
            cases_payload = {
                "schemaVersion": 1,
                "cases": [{"id": sensitive_case_id, "formula": "=1+1", "outputCell": "C1", "inputs": []}],
            }
            cases_path.write_text(
                json.dumps(cases_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            expected_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": sensitive_case_id, "result": {"t": "n", "v": 2}}],
            }
            expected_path.write_text(
                json.dumps(expected_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            actual_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 1},
                "results": [{"caseId": sensitive_case_id, "result": {"t": "n", "v": 3}}],
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
                    "--privacy-mode",
                    "private",
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

            expected_hash = hashlib.sha256(sensitive_case_id.encode("utf-8")).hexdigest()
            self.assertEqual(mismatches[0].get("caseId"), f"sha256={expected_hash}")
            self.assertNotIn(sensitive_case_id, json.dumps(report))

    def test_privacy_mode_hashes_non_standard_missing_function_names(self) -> None:
        compare_py = Path(__file__).resolve().parents[1] / "compare.py"
        self.assertTrue(compare_py.is_file(), f"compare.py not found at {compare_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            cases_path = tmp_path / "cases.json"
            expected_path = tmp_path / "expected.json"
            actual_path = tmp_path / "actual.json"
            report_path = tmp_path / "report.json"

            # Two cases that both produce #NAME? in the actual results. One is a built-in function
            # (SUM) that should remain readable; the other is a synthetic UDF-style name that should
            # be hashed in privacy mode.
            cases_payload = {
                "schemaVersion": 1,
                "cases": [
                    {"id": "case-sum", "formula": "=SUM(1,2)", "outputCell": "C1", "inputs": []},
                    {
                        "id": "case-udf",
                        "formula": "=CORP.ADDIN.FOO(1)",
                        "outputCell": "C1",
                        "inputs": [],
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
                "source": {"kind": "excel"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "n", "v": 3}},
                    {"caseId": "case-udf", "result": {"t": "n", "v": 1}},
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
                "source": {"kind": "engine"},
                "caseSet": {"path": str(cases_path), "count": 2},
                "results": [
                    {"caseId": "case-sum", "result": {"t": "e", "v": "#NAME?"}},
                    {"caseId": "case-udf", "result": {"t": "e", "v": "#NAME?"}},
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
                    "--privacy-mode",
                    "private",
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
            top_missing = report.get("summary", {}).get("topMissingFunctions")
            self.assertIsInstance(top_missing, list)

            names = {row.get("name") for row in top_missing if isinstance(row, dict)}
            self.assertIn("SUM", names)
            udf_name = "CORP.ADDIN.FOO"
            self.assertNotIn(udf_name, names)
            udf_hash = hashlib.sha256(udf_name.encode("utf-8")).hexdigest()
            self.assertIn(f"sha256={udf_hash}", names)

            mismatches = report.get("mismatches")
            self.assertIsInstance(mismatches, list)
            by_id = {m.get("caseId"): m for m in mismatches if isinstance(m, dict)}
            udf_formula = by_id.get("case-udf", {}).get("formula")
            self.assertIsInstance(udf_formula, str)
            self.assertNotIn(udf_name, udf_formula)
            self.assertIn(f"sha256={udf_hash}", udf_formula)


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
