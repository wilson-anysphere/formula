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

        # Source label should be stable across Excel updates so we don't churn diffs when only the
        # Office build number changes.
        self.assertIn(
            'source = "Microsoft Excel ($LocaleId) function name translations via Range.Formula/FormulaLocal round-trip',
            text,
        )

        # Sentinel translations must match what we expect for a correctly configured Excel locale.
        # (These are used only for warnings, but keeping them accurate prevents confusing output.)
        self.assertIn('SUM = "SUMME"', text)
        self.assertIn('IF = "WENN"', text)
        self.assertIn('TRUE = "WAHR"', text)
        self.assertIn('FALSE = "FALSCH"', text)


if __name__ == "__main__":
    unittest.main()
