from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools.corpus.triage import _resolve_triage_input_dir


class TriageInputScopeTests(unittest.TestCase):
    def test_auto_prefers_sanitized_subdir_when_present(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sanitized = root / "sanitized"
            originals = root / "originals"
            sanitized.mkdir(parents=True)
            originals.mkdir(parents=True)

            # The file content does not matter for scope resolution; these are just "workbook-like"
            # blobs for a realistic directory layout.
            (sanitized / "book.xlsx").write_bytes(b"dummy")
            (originals / "book.xlsx.enc").write_bytes(b"dummy")

            chosen = _resolve_triage_input_dir(root, "auto")
            self.assertEqual(chosen, sanitized)


if __name__ == "__main__":
    unittest.main()

