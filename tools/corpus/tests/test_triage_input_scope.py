from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

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

    def test_main_auto_scans_only_sanitized_tree_when_present(self) -> None:
        import tools.corpus.triage as triage_mod

        observed: dict[str, object] = {}

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
                observed["paths"] = list(paths)
                # Return a minimal report per path so main() can write JSON artifacts.
                return [
                    {
                        "display_name": p.name,
                        "sha256": "0" * 64,
                        "result": {"open_ok": True, "round_trip_ok": True},
                    }
                    for p in paths
                ]

            triage_mod._triage_paths = _fake_triage_paths  # type: ignore[assignment]

            with tempfile.TemporaryDirectory() as tmp:
                corpus_dir = Path(tmp) / "corpus"
                sanitized = corpus_dir / "sanitized"
                originals = corpus_dir / "originals"
                sanitized.mkdir(parents=True)
                originals.mkdir(parents=True)
                (sanitized / "a.xlsx").write_bytes(b"dummy-xlsx")
                (originals / "b.xlsx.enc").write_bytes(b"dummy-enc")

                out_dir = Path(tmp) / "out"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.triage",
                        "--corpus-dir",
                        str(corpus_dir),
                        "--out-dir",
                        str(out_dir),
                    ]
                    buf = io.StringIO()
                    with mock.patch("sys.stdout", buf):
                        rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                self.assertIn("paths", observed)
                paths = observed["paths"]
                self.assertIsInstance(paths, list)
                rel_paths = [p.relative_to(corpus_dir).as_posix() for p in paths]
                self.assertEqual(rel_paths, ["sanitized/a.xlsx"])

                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertEqual(index["input_scope"], "auto")
                self.assertEqual(index["input_dir"], str(sanitized))
                self.assertEqual(index["report_count"], 1)
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]


if __name__ == "__main__":
    unittest.main()
