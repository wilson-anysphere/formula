from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class PinDatasetTests(unittest.TestCase):
    def test_refuses_unknown_excel_metadata(self) -> None:
        pin_py = Path(__file__).resolve().parents[1] / "pin_dataset.py"
        self.assertTrue(pin_py.is_file(), f"pin_dataset.py not found at {pin_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            dataset_path = tmp / "dataset.json"
            pinned_path = tmp / "pinned.json"
            versioned_dir = tmp / "versioned"

            dataset_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {
                    "kind": "excel",
                    "version": "unknown",
                    "build": "unknown",
                    "operatingSystem": "unknown",
                },
                "caseSet": {"path": "cases.json", "sha256": "deadbeef" * 8, "count": 0},
                "results": [],
            }
            dataset_path.write_text(
                json.dumps(dataset_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(pin_py),
                    "--dataset",
                    str(dataset_path),
                    "--pinned",
                    str(pinned_path),
                    "--versioned-dir",
                    str(versioned_dir),
                ],
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(proc.returncode, 0)
            self.assertIn(
                "Refusing to pin dataset missing Excel metadata",
                proc.stdout + proc.stderr,
            )

    def test_allows_formula_engine_synthetic_dataset(self) -> None:
        pin_py = Path(__file__).resolve().parents[1] / "pin_dataset.py"
        self.assertTrue(pin_py.is_file(), f"pin_dataset.py not found at {pin_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            dataset_path = tmp / "dataset.json"
            pinned_path = tmp / "pinned.json"
            versioned_dir = tmp / "versioned"

            dataset_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {
                    "kind": "formula-engine",
                    "version": "0.0.0-test",
                    "os": "linux",
                    "arch": "x86_64",
                    "caseSet": "unit-test",
                },
                "caseSet": {"path": "cases.json", "sha256": "deadbeef" * 8, "count": 0},
                "results": [],
            }
            dataset_path.write_text(
                json.dumps(dataset_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(pin_py),
                    "--dataset",
                    str(dataset_path),
                    "--pinned",
                    str(pinned_path),
                    "--versioned-dir",
                    str(versioned_dir),
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(f"pin_dataset.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}")

            pinned = json.loads(pinned_path.read_text(encoding="utf-8"))
            self.assertEqual(pinned.get("source", {}).get("kind"), "excel")
            self.assertEqual(pinned.get("source", {}).get("version"), "unknown")
            self.assertIn("syntheticSource", pinned.get("source", {}))
            self.assertEqual(pinned.get("source", {}).get("syntheticSource", {}).get("kind"), "formula-engine")

            expected_versioned = versioned_dir / "excel-unknown-build-unknown-cases-deadbeef.json"
            self.assertTrue(
                expected_versioned.is_file(),
                f"expected versioned dataset at {expected_versioned} (got: {list(versioned_dir.glob('*.json'))})",
            )

    def test_writes_pinned_and_versioned_copy(self) -> None:
        pin_py = Path(__file__).resolve().parents[1] / "pin_dataset.py"
        self.assertTrue(pin_py.is_file(), f"pin_dataset.py not found at {pin_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            dataset_path = tmp / "dataset.json"
            pinned_path = tmp / "pinned.json"
            versioned_dir = tmp / "versioned"

            dataset_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                "source": {
                    "kind": "excel",
                    "version": "16.0",
                    "build": "12345.67890",
                    "operatingSystem": "Windows",
                },
                "caseSet": {"path": "cases.json", "sha256": "deadbeef" * 8, "count": 0},
                "results": [],
            }
            dataset_text = json.dumps(dataset_payload, ensure_ascii=False, indent=2) + "\n"
            dataset_path.write_text(dataset_text, encoding="utf-8", newline="\n")

            proc = subprocess.run(
                [
                    sys.executable,
                    str(pin_py),
                    "--dataset",
                    str(dataset_path),
                    "--pinned",
                    str(pinned_path),
                    "--versioned-dir",
                    str(versioned_dir),
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(f"pin_dataset.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}")

            self.assertTrue(pinned_path.is_file())
            self.assertEqual(pinned_path.read_text(encoding="utf-8"), dataset_text)

            expected_versioned = versioned_dir / "excel-16.0-build-12345.67890-cases-deadbeef.json"
            self.assertTrue(
                expected_versioned.is_file(),
                f"expected versioned dataset at {expected_versioned} (got: {list(versioned_dir.glob('*.json'))})",
            )
            self.assertEqual(expected_versioned.read_text(encoding="utf-8"), dataset_text)

    def test_dry_run_does_not_write_files(self) -> None:
        pin_py = Path(__file__).resolve().parents[1] / "pin_dataset.py"
        self.assertTrue(pin_py.is_file(), f"pin_dataset.py not found at {pin_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            dataset_path = tmp / "dataset.json"
            pinned_path = tmp / "pinned.json"
            versioned_dir = tmp / "versioned"

            dataset_payload = {
                "schemaVersion": 1,
                "generatedAt": "unit-test",
                # Use the synthetic baseline path so this test is platform-independent.
                "source": {
                    "kind": "formula-engine",
                    "version": "0.0.0-test",
                    "os": "linux",
                    "arch": "x86_64",
                    "caseSet": "unit-test",
                },
                "caseSet": {"path": "cases.json", "sha256": "deadbeef" * 8, "count": 0},
                "results": [],
            }
            dataset_path.write_text(
                json.dumps(dataset_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(pin_py),
                    "--dataset",
                    str(dataset_path),
                    "--pinned",
                    str(pinned_path),
                    "--versioned-dir",
                    str(versioned_dir),
                    "--dry-run",
                ],
                capture_output=True,
                text=True,
            )
            if proc.returncode != 0:
                self.fail(f"pin_dataset.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}")

            # Dry-run should not create outputs.
            self.assertFalse(pinned_path.exists())
            self.assertFalse(versioned_dir.exists())
            self.assertIn("Dry run: pin_dataset", proc.stdout)


if __name__ == "__main__":
    unittest.main()
