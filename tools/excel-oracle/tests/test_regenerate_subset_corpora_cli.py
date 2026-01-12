from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class RegenerateSubsetCorporaCliTests(unittest.TestCase):
    def test_help_does_not_write_files(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "regenerate_subset_corpora.py"
        self.assertTrue(script.is_file(), f"script not found at {script}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            out_dir = tmp / "out"
            out_dir.mkdir(parents=True, exist_ok=True)

            proc = subprocess.run(
                [sys.executable, str(script), "--help", "--out-dir", str(out_dir)],
                cwd=str(repo_root),
                capture_output=True,
                text=True,
            )
            self.assertEqual(proc.returncode, 0)
            self.assertTrue(proc.stdout, "expected --help to print usage")
            self.assertEqual(list(out_dir.iterdir()), [], "--help should not create any files")

    def test_check_mode_detects_drift(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "regenerate_subset_corpora.py"
        self.assertTrue(script.is_file(), f"script not found at {script}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            cases_path = tmp / "cases.json"
            out_dir = tmp / "out"

            cases_payload = {
                "schemaVersion": 1,
                "caseSet": "test",
                "defaultSheet": "Sheet1",
                "cases": [
                    {
                        "id": "v1",
                        "formula": "=1",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["odd_coupon_validation"],
                    },
                    {
                        "id": "b1",
                        "formula": "=2",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["odd_coupon", "boundary"],
                    },
                    {
                        "id": "l1",
                        "formula": "=3",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["odd_coupon", "long_stub"],
                    },
                    {
                        "id": "i1",
                        "formula": "=4",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["odd_coupon", "invalid_schedule"],
                    },
                    # Unrelated case should not appear in any subset.
                    {
                        "id": "x1",
                        "formula": "=5",
                        "outputCell": "C1",
                        "inputs": [],
                        "tags": ["arith"],
                    },
                ],
            }
            cases_path.write_text(
                json.dumps(cases_payload, indent=2, sort_keys=False) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            # Generate subset corpora.
            subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--cases",
                    str(cases_path),
                    "--out-dir",
                    str(out_dir),
                ],
                cwd=str(repo_root),
                check=True,
                capture_output=True,
                text=True,
            )

            # Check should pass when corpora match.
            proc = subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--cases",
                    str(cases_path),
                    "--out-dir",
                    str(out_dir),
                    "--check",
                ],
                cwd=str(repo_root),
                capture_output=True,
                text=True,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)

            # Mutate a subset file to simulate drift.
            boundary_path = out_dir / "odd_coupon_boundary_cases.json"
            self.assertTrue(boundary_path.is_file())
            boundary_payload = json.loads(boundary_path.read_text(encoding="utf-8"))
            boundary_payload["cases"] = []
            boundary_path.write_text(
                json.dumps(boundary_payload, indent=2, sort_keys=False) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--cases",
                    str(cases_path),
                    "--out-dir",
                    str(out_dir),
                    "--check",
                ],
                cwd=str(repo_root),
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(proc.returncode, 0, "expected --check to fail on drift")
            self.assertIn("Subset corpora are out of date", proc.stderr + proc.stdout)


if __name__ == "__main__":
    unittest.main()
