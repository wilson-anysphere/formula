from __future__ import annotations

import unittest
from pathlib import Path


class RunExcelOracleDryRunParamTests(unittest.TestCase):
    def test_script_exposes_dry_run_switch(self) -> None:
        """The Windows-only PowerShell runner should have a discoverable -DryRun mode."""

        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "run-excel-oracle.ps1"
        self.assertTrue(script.is_file(), f"run-excel-oracle.ps1 not found at {script}")

        text = script.read_text(encoding="utf-8")
        self.assertIn(".PARAMETER DryRun", text)
        self.assertIn("[switch]$DryRun", text)
        self.assertIn("if ($DryRun)", text)


if __name__ == "__main__":
    unittest.main()

