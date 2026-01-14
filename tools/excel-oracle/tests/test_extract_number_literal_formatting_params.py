from __future__ import annotations

import unittest
from pathlib import Path


class ExtractNumberLiteralFormattingScriptTests(unittest.TestCase):
    def test_script_exposes_expected_parameters(self) -> None:
        """The PowerShell numeric-literal probe should expose stable CLI parameters (no Excel required)."""

        repo_root = Path(__file__).resolve().parents[3]
        script = (
            repo_root
            / "tools"
            / "excel-oracle"
            / "extract-number-literal-formatting.ps1"
        )
        self.assertTrue(
            script.is_file(),
            f"extract-number-literal-formatting.ps1 not found at {script}",
        )

        text = script.read_text(encoding="utf-8")

        # Comment-based help for discoverability.
        self.assertIn(".PARAMETER LocaleId", text)
        self.assertIn(".PARAMETER OutPath", text)
        self.assertIn(".PARAMETER Visible", text)

        # Parameter block (CmdletBinding param()).
        self.assertIn("[string]$LocaleId", text)
        self.assertIn("[string]$OutPath", text)
        self.assertIn("[switch]$Visible", text)

        # Core mechanism should round-trip through FormulaLocal and prefer Formula2.
        self.assertIn(".FormulaLocal", text)
        self.assertIn(".Formula2", text)

        # Script should probe the sentinel formulas we care about.
        self.assertIn("=SUM(1234.56,0.5)", text)
        self.assertIn("=SUM(1234567.89,0.5)", text)
        self.assertIn("=SUM(1000,0)", text)
        self.assertIn("=SUM(0001,0)", text)
        self.assertIn("=SUM(1E3,0)", text)


if __name__ == "__main__":
    unittest.main()

