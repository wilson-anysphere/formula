from __future__ import annotations

import unittest
from pathlib import Path


class ExtractErrorLiteralsScriptTests(unittest.TestCase):
    def test_script_exposes_expected_parameters(self) -> None:
        """The PowerShell error-literal extractor should expose stable CLI parameters (no Excel required)."""

        repo_root = Path(__file__).resolve().parents[3]
        script = repo_root / "tools" / "excel-oracle" / "extract-error-literals.ps1"
        self.assertTrue(script.is_file(), f"extract-error-literals.ps1 not found at {script}")

        text = script.read_text(encoding="utf-8")

        # Comment-based help for discoverability.
        self.assertTrue(
            (".PARAMETER Locale" in text) or (".PARAMETER LocaleId" in text),
            "Comment-based help must document a Locale/LocaleId parameter",
        )
        self.assertIn(".PARAMETER OutPath", text)
        self.assertIn(".PARAMETER Visible", text)

        # Parameter block (CmdletBinding param()).
        # Accept either `$Locale` or `$LocaleId` (some scripts use LocaleId for consistency).
        self.assertTrue(
            ("[string]$Locale" in text) or ("[string]$LocaleId" in text),
            "Script must expose a string Locale/LocaleId parameter in its param() block",
        )
        self.assertIn("[string]$OutPath", text)
        self.assertIn("[switch]$Visible", text)

        # The extractor should stay tied to the canonical error literal source and the core
        # COM round-tripping mechanism for localization.
        self.assertIn("crates/formula-engine/src/value/mod.rs", text)
        self.assertIn("ErrorKind::as_code", text)
        self.assertIn(".FormulaLocal", text)
        self.assertIn(".Formula2", text)


if __name__ == "__main__":
    unittest.main()

