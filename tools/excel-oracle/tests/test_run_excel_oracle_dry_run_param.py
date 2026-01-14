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

    def test_script_stabilizes_case_set_path_for_uri_like_inputs(self) -> None:
        """Get-StableCaseSetPath should treat file://... style paths as absolute (privacy-safe)."""

        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "run-excel-oracle.ps1"
        self.assertTrue(script.is_file(), f"run-excel-oracle.ps1 not found at {script}")

        text = script.read_text(encoding="utf-8")
        self.assertIn("Get-StableCaseSetPath", text)
        # Heuristic: ensure the path normalization treats URI schemes as absolute, not "relative".
        self.assertIn("Treat URI-like paths", text)
        self.assertIn("^[A-Za-z][A-Za-z0-9+.-]*:", text)


if __name__ == "__main__":
    unittest.main()
