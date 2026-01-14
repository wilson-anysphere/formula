from __future__ import annotations

import unittest
from pathlib import Path


class ExtractStructuredReferenceKeywordsScriptTests(unittest.TestCase):
    def test_script_exposes_expected_parameters(self) -> None:
        """The PowerShell structured-ref probe should expose stable CLI parameters (no Excel required)."""

        repo_root = Path(__file__).resolve().parents[3]
        script = (
            repo_root
            / "tools"
            / "excel-oracle"
            / "extract-structured-reference-keywords.ps1"
        )
        self.assertTrue(
            script.is_file(),
            f"extract-structured-reference-keywords.ps1 not found at {script}",
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

        # Should create a table to ensure structured refs resolve.
        self.assertIn("ListObjects.Add", text)
        self.assertIn('$table.Name = "Table1"', text)


if __name__ == "__main__":
    unittest.main()

