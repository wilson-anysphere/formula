from __future__ import annotations

import unittest
from pathlib import Path


class ExtractFunctionTranslationsScriptTests(unittest.TestCase):
    def test_script_exposes_expected_parameters(self) -> None:
        """The PowerShell extractor should expose stable CLI parameters (no Excel required)."""

        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "extract-function-translations.ps1"
        self.assertTrue(
            script.is_file(), f"extract-function-translations.ps1 not found at {script}"
        )

        text = script.read_text(encoding="utf-8")

        # Comment-based help for discoverability.
        self.assertIn(".PARAMETER LocaleId", text)
        self.assertIn(".PARAMETER OutPath", text)
        self.assertIn(".PARAMETER Visible", text)
        self.assertIn(".PARAMETER MaxFunctions", text)

        # Parameter block (CmdletBinding param()).
        self.assertIn("[string]$LocaleId", text)
        self.assertIn("[string]$OutPath", text)
        self.assertIn("[switch]$Visible", text)
        self.assertIn("[int]$MaxFunctions", text)

        # Core extraction mechanism should round-trip through FormulaLocal.
        self.assertIn(".FormulaLocal", text)
        self.assertIn(".Formula2", text)
        self.assertIn("functionCatalog.json", text)
        self.assertIn('"shared"', text)


if __name__ == "__main__":
    unittest.main()
