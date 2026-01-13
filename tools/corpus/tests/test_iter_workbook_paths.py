from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools.corpus.util import iter_workbook_paths


class IterWorkbookPathsTests(unittest.TestCase):
    def test_iter_workbook_paths_excludes_xlsb_by_default(self) -> None:
        with tempfile.TemporaryDirectory(prefix="corpus-iter-paths-") as td:
            corpus_dir = Path(td)
            (corpus_dir / "a.xlsx").write_text("", encoding="utf-8")
            (corpus_dir / "b.xlsm").write_text("", encoding="utf-8")
            (corpus_dir / "c.xlsb").write_text("", encoding="utf-8")
            (corpus_dir / "d.xlsb.b64").write_text("", encoding="utf-8")
            (corpus_dir / "e.xlsb.enc").write_text("", encoding="utf-8")

            paths = [p.name for p in iter_workbook_paths(corpus_dir)]
            self.assertEqual(paths, ["a.xlsx", "b.xlsm"])

    def test_iter_workbook_paths_includes_xlsb_when_enabled(self) -> None:
        with tempfile.TemporaryDirectory(prefix="corpus-iter-paths-") as td:
            corpus_dir = Path(td)
            (corpus_dir / "a.xlsx").write_text("", encoding="utf-8")
            (corpus_dir / "b.xlsm").write_text("", encoding="utf-8")
            (corpus_dir / "c.xlsb").write_text("", encoding="utf-8")
            (corpus_dir / "d.xlsb.b64").write_text("", encoding="utf-8")
            (corpus_dir / "e.xlsb.enc").write_text("", encoding="utf-8")

            paths = [p.name for p in iter_workbook_paths(corpus_dir, include_xlsb=True)]
            self.assertEqual(paths, ["a.xlsx", "b.xlsm", "c.xlsb", "d.xlsb.b64", "e.xlsb.enc"])


if __name__ == "__main__":
    unittest.main()

