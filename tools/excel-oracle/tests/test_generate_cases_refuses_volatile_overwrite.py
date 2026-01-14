from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path


class GenerateCasesVolatileOverwriteGuardTests(unittest.TestCase):
    def test_include_volatile_refuses_to_overwrite_committed_cases_json(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        gen_py = repo_root / "tools/excel-oracle/generate_cases.py"
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"

        self.assertTrue(gen_py.is_file(), f"generate_cases.py not found at {gen_py}")
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")

        before = cases_path.read_bytes()

        out_args = [
            # Use the canonical relative path (what a dev is likely to copy/paste).
            "tests/compatibility/excel-oracle/cases.json",
            # Also reject the absolute path (what a script/wrapper might provide).
            str(cases_path.resolve()),
        ]

        for out in out_args:
            proc = subprocess.run(
                [
                    sys.executable,
                    str(gen_py),
                    "--include-volatile",
                    "--out",
                    out,
                ],
                capture_output=True,
                text=True,
                cwd=str(repo_root),
            )

            self.assertNotEqual(proc.returncode, 0)
            self.assertIn("Refusing to write a volatile debug corpus", proc.stdout + proc.stderr)

            after = cases_path.read_bytes()
            self.assertEqual(
                after,
                before,
                "cases.json should not be modified when overwrite is refused",
            )


if __name__ == "__main__":
    unittest.main()
