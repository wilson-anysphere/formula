from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import sys
import tempfile
import unittest
import io
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock


class UpdatePinnedDatasetTests(unittest.TestCase):
    @contextmanager
    def _patched_argv(self, argv: list[str]):
        old = sys.argv[:]
        sys.argv = argv
        try:
            yield
        finally:
            sys.argv = old

    def _load_update_module(self):
        tool = Path(__file__).resolve().parents[1] / "update_pinned_dataset.py"
        self.assertTrue(tool.is_file(), f"update_pinned_dataset.py not found at {tool}")

        spec = importlib.util.spec_from_file_location("excel_oracle_update_pinned_dataset", tool)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_merges_results_without_running_engine(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                    {"id": "case2", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")
            cases_sha = hashlib.sha256(cases_path.read_bytes()).hexdigest()

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {"kind": "excel", "version": "unknown", "build": "unknown", "operatingSystem": "unknown"},
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 1},
                "results": [{"caseId": "case1"}],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            merge_path = tmp / "merge.json"
            merge_payload = {"schemaVersion": 1, "results": [{"caseId": "case2"}]}
            merge_path.write_text(json.dumps(merge_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            missing_before, missing_after = update.update_pinned_dataset(
                cases_path=cases_path,
                pinned_path=pinned_path,
                merge_results_paths=[merge_path],
                engine_bin=None,
                run_engine_for_missing=False,
            )
            self.assertEqual(missing_before, 1)
            self.assertEqual(missing_after, 0)

            pinned_updated = json.loads(pinned_path.read_text(encoding="utf-8"))
            self.assertEqual(pinned_updated["caseSet"]["sha256"], cases_sha)
            self.assertEqual(pinned_updated["caseSet"]["count"], 2)
            result_ids = {r.get("caseId") for r in pinned_updated.get("results", []) if isinstance(r, dict)}
            self.assertEqual(result_ids, {"case1", "case2"})

            versioned_dir = tmp / "versioned"
            versioned_path = update.write_versioned_copy(pinned_path=pinned_path, versioned_dir=versioned_dir)
            self.assertTrue(versioned_path.is_file())
            self.assertEqual(
                versioned_path.name,
                f"excel-unknown-build-unknown-cases-{cases_sha[:8]}.json",
            )
            versioned_payload = json.loads(versioned_path.read_text(encoding="utf-8"))
            self.assertEqual(versioned_payload["caseSet"]["sha256"], cases_sha)

    def test_cli_writes_versioned_copy_by_default(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                    {"id": "case2", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")
            cases_sha = hashlib.sha256(cases_path.read_bytes()).hexdigest()

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {
                    "kind": "excel",
                    "version": "unknown",
                    "build": "unknown",
                    "operatingSystem": "unknown",
                },
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 2},
                "results": [{"caseId": "case1"}, {"caseId": "case2"}],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            versioned_dir = tmp / "versioned"
            argv = [
                str(Path(update.__file__)),
                "--cases",
                str(cases_path),
                "--pinned",
                str(pinned_path),
                "--no-engine",
                "--versioned-dir",
                str(versioned_dir),
            ]
            with self._patched_argv(argv):
                buf = io.StringIO()
                with redirect_stdout(buf), redirect_stderr(buf):
                    rc = update.main()
            self.assertEqual(rc, 0)

            expected_versioned = versioned_dir / f"excel-unknown-build-unknown-cases-{cases_sha[:8]}.json"
            self.assertTrue(expected_versioned.is_file())

            pinned_updated = json.loads(pinned_path.read_text(encoding="utf-8"))
            versioned_payload = json.loads(expected_versioned.read_text(encoding="utf-8"))
            self.assertEqual(versioned_payload, pinned_updated)

    def test_cli_no_versioned_skips_copy(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                    {"id": "case2", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {
                    "kind": "excel",
                    "version": "unknown",
                    "build": "unknown",
                    "operatingSystem": "unknown",
                },
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 2},
                "results": [{"caseId": "case1"}, {"caseId": "case2"}],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            versioned_dir = tmp / "versioned"
            argv = [
                str(Path(update.__file__)),
                "--cases",
                str(cases_path),
                "--pinned",
                str(pinned_path),
                "--no-engine",
                "--no-versioned",
                "--versioned-dir",
                str(versioned_dir),
            ]
            with self._patched_argv(argv):
                buf = io.StringIO()
                with redirect_stdout(buf), redirect_stderr(buf):
                    rc = update.main()
            self.assertEqual(rc, 0)
            self.assertFalse(versioned_dir.exists(), "--no-versioned should avoid creating the dir")

    def test_overwrite_existing_replaces_case_results(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                    {"id": "case2", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")
            cases_sha = hashlib.sha256(cases_path.read_bytes()).hexdigest()

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {"kind": "excel", "version": "unknown", "build": "unknown", "operatingSystem": "unknown"},
                "caseSet": {"path": "cases.json", "sha256": cases_sha, "count": 2},
                "results": [
                    {"caseId": "case1", "result": {"t": "n", "v": 2}},
                    {"caseId": "case2", "result": {"t": "n", "v": 4}},
                ],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            merge_path = tmp / "merge.json"
            merge_payload = {"schemaVersion": 1, "results": [{"caseId": "case1", "result": {"t": "n", "v": 3}}]}
            merge_path.write_text(json.dumps(merge_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            missing_before, missing_after = update.update_pinned_dataset(
                cases_path=cases_path,
                pinned_path=pinned_path,
                merge_results_paths=[merge_path],
                engine_bin=None,
                run_engine_for_missing=False,
                overwrite_existing=True,
            )
            self.assertEqual(missing_before, 0)
            self.assertEqual(missing_after, 0)

            pinned_updated = json.loads(pinned_path.read_text(encoding="utf-8"))
            results = pinned_updated.get("results", [])
            self.assertIsInstance(results, list)
            by_id = {r.get("caseId"): r for r in results if isinstance(r, dict)}
            self.assertEqual(by_id["case1"]["result"]["v"], 3)
            self.assertEqual(by_id["case2"]["result"]["v"], 4)

    def test_records_real_excel_patch_metadata(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")
            cases_sha = hashlib.sha256(cases_path.read_bytes()).hexdigest()

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {
                    "kind": "excel",
                    "version": "unknown",
                    "build": "unknown",
                    "operatingSystem": "unknown",
                    "note": "Synthetic CI baseline (not generated by Excel).",
                    "syntheticSource": {"kind": "formula-engine"},
                },
                "caseSet": {"path": "cases.json", "sha256": cases_sha, "count": 1},
                "results": [{"caseId": "case1", "result": {"t": "n", "v": 2}}],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            # Simulate a merge-results payload produced by real Excel (no syntheticSource metadata).
            merge_path = tmp / "merge.json"
            merge_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-02T00:00:00Z",
                "source": {"kind": "excel", "version": "16.0", "build": "12345", "operatingSystem": "Windows"},
                "caseSet": {"path": "tools/excel-oracle/odd_coupon_boundary_cases.json", "sha256": "deadbeef", "count": 1},
                "results": [{"caseId": "case1", "result": {"t": "n", "v": 3}}],
            }
            merge_path.write_text(json.dumps(merge_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            update.update_pinned_dataset(
                cases_path=cases_path,
                pinned_path=pinned_path,
                merge_results_paths=[merge_path],
                engine_bin=None,
                run_engine_for_missing=False,
                overwrite_existing=True,
            )

            pinned_updated = json.loads(pinned_path.read_text(encoding="utf-8"))
            source = pinned_updated.get("source", {})
            self.assertIsInstance(source, dict)
            patches = source.get("patches")
            self.assertIsInstance(patches, list)
            self.assertEqual(len(patches), 1)
            patch = patches[0]
            self.assertEqual(patch.get("version"), "16.0")
            self.assertEqual(patch.get("build"), "12345")
            self.assertEqual(patch.get("operatingSystem"), "Windows")
            self.assertEqual(patch.get("applied", {}).get("overwritten"), 1)
            self.assertEqual(patch.get("caseSet", {}).get("path"), "tools/excel-oracle/odd_coupon_boundary_cases.json")
            self.assertIn("real Excel patches", source.get("note", ""))

    def test_fails_if_missing_and_no_engine(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [{"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []}],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {"kind": "excel", "version": "unknown", "build": "unknown", "operatingSystem": "unknown"},
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 0},
                "results": [],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            with self.assertRaises(SystemExit):
                update.update_pinned_dataset(
                    cases_path=cases_path,
                    pinned_path=pinned_path,
                    merge_results_paths=[],
                    engine_bin=None,
                    run_engine_for_missing=False,
                )

    def test_refuses_to_fill_real_excel_dataset_with_engine(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []},
                    {"id": "case2", "formula": "=2+2", "outputCell": "C1", "inputs": [], "tags": []},
                ],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            pinned_path = tmp / "pinned.json"
            # Simulate a dataset pinned from real Excel: source.kind=excel and no syntheticSource metadata.
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                "source": {"kind": "excel", "version": "16.0", "build": "12345", "operatingSystem": "Windows"},
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 1},
                "results": [{"caseId": "case1"}],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            # By default the updater should refuse to fill missing cases by running the engine when the
            # dataset looks like real Excel.
            with self.assertRaises(SystemExit):
                update.update_pinned_dataset(
                    cases_path=cases_path,
                    pinned_path=pinned_path,
                    merge_results_paths=[],
                    engine_bin=None,
                    run_engine_for_missing=True,
                )

    def test_cli_dry_run_does_not_write_or_run_engine(self) -> None:
        update = self._load_update_module()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)

            cases_path = tmp / "cases.json"
            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [{"id": "case1", "formula": "=1+1", "outputCell": "C1", "inputs": [], "tags": []}],
            }
            cases_path.write_text(json.dumps(cases_payload, indent=2) + "\n", encoding="utf-8", newline="\n")

            pinned_path = tmp / "pinned.json"
            pinned_payload = {
                "schemaVersion": 1,
                "generatedAt": "2026-01-01T00:00:00Z",
                # Mark as a synthetic baseline so the updater would normally be willing to run the engine.
                "source": {
                    "kind": "excel",
                    "version": "unknown",
                    "build": "unknown",
                    "operatingSystem": "unknown",
                    "syntheticSource": {"kind": "formula-engine"},
                },
                "caseSet": {"path": "cases.json", "sha256": "old", "count": 0},
                "results": [],
            }
            pinned_path.write_text(json.dumps(pinned_payload, indent=2) + "\n", encoding="utf-8", newline="\n")
            pinned_before = pinned_path.read_text(encoding="utf-8")

            versioned_dir = tmp / "versioned"
            argv = [
                str(Path(update.__file__)),
                "--cases",
                str(cases_path),
                "--pinned",
                str(pinned_path),
                "--versioned-dir",
                str(versioned_dir),
                "--dry-run",
            ]
            with self._patched_argv(argv):
                buf = io.StringIO()
                with mock.patch.object(update.subprocess, "run") as run_mock, redirect_stdout(buf), redirect_stderr(buf):
                    rc = update.main()
            self.assertEqual(rc, 0)
            run_mock.assert_not_called()

            # Dry-run should not modify the pinned dataset file or write versioned copies.
            self.assertEqual(pinned_path.read_text(encoding="utf-8"), pinned_before)
            self.assertFalse(versioned_dir.exists())


if __name__ == "__main__":
    unittest.main()
