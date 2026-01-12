from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class CasesSyncTests(unittest.TestCase):
    def test_cases_json_matches_generator(self) -> None:
        """
        Ensure `tests/compatibility/excel-oracle/cases.json` is generated from
        `tools/excel-oracle/generate_cases.py`.

        This protects against a common integration failure mode when new deterministic
        functions are added:
        - update functionCatalog + generator
        - forget to regenerate/commit cases.json
        """

        repo_root = Path(__file__).resolve().parents[3]
        gen_py = repo_root / "tools/excel-oracle/generate_cases.py"
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"

        self.assertTrue(gen_py.is_file(), f"generate_cases.py not found at {gen_py}")
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_cases = Path(tmp_dir) / "cases.json"

            proc = subprocess.run(
                [sys.executable, str(gen_py), "--out", str(tmp_cases)],
                capture_output=True,
                text=True,
                cwd=str(repo_root),
            )
            if proc.returncode != 0:
                self.fail(
                    f"generate_cases.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            generated = json.loads(tmp_cases.read_text(encoding="utf-8"))
            committed = json.loads(cases_path.read_text(encoding="utf-8"))

            self.assertEqual(
                committed,
                generated,
                "tests/compatibility/excel-oracle/cases.json is out of sync with tools/excel-oracle/generate_cases.py.\n"
                "Regenerate it with:\n"
                "  python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json",
            )


if __name__ == "__main__":
    unittest.main()

