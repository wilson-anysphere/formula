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

    def test_auto_falls_back_to_corpus_dir_when_sanitized_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            chosen = _resolve_triage_input_dir(root, "auto")
            self.assertEqual(chosen, root)

    def test_sanitized_scope_requires_sanitized_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with self.assertRaises(ValueError):
                _resolve_triage_input_dir(root, "sanitized")

    def test_originals_scope_requires_originals_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with self.assertRaises(ValueError):
                _resolve_triage_input_dir(root, "originals")

    def test_all_scope_uses_corpus_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            chosen = _resolve_triage_input_dir(root, "all")
            self.assertEqual(chosen, root)

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

    def test_main_private_mode_hashes_index_paths(self) -> None:
        import tools.corpus.triage as triage_mod
        from tools.corpus.util import sha256_hex

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
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
                sanitized.mkdir(parents=True)
                (sanitized / "a.xlsx").write_bytes(b"dummy-xlsx")

                out_dir = Path(tmp) / "out"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.triage",
                        "--corpus-dir",
                        str(corpus_dir),
                        "--out-dir",
                        str(out_dir),
                        "--privacy-mode",
                        "private",
                    ]
                    with mock.patch("sys.stdout", new=io.StringIO()):
                        rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertEqual(
                    index["corpus_dir"], f"sha256={sha256_hex(str(corpus_dir).encode('utf-8'))}"
                )
                self.assertEqual(
                    index["input_dir"], f"sha256={sha256_hex(str(sanitized).encode('utf-8'))}"
                )
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]

    def test_expectations_are_skipped_when_input_is_scoped(self) -> None:
        import tools.corpus.triage as triage_mod

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        original_compare_expectations = triage_mod._compare_expectations
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
                return [
                    {
                        "display_name": p.name,
                        "sha256": "0" * 64,
                        "result": {"open_ok": True, "round_trip_ok": True},
                    }
                    for p in paths
                ]

            triage_mod._triage_paths = _fake_triage_paths  # type: ignore[assignment]

            def _boom(*_args, **_kwargs):  # type: ignore[no-untyped-def]
                raise AssertionError("_compare_expectations should not run for scoped inputs")

            triage_mod._compare_expectations = _boom  # type: ignore[assignment]

            with tempfile.TemporaryDirectory() as tmp:
                corpus_dir = Path(tmp) / "corpus"
                sanitized = corpus_dir / "sanitized"
                originals = corpus_dir / "originals"
                sanitized.mkdir(parents=True)
                originals.mkdir(parents=True)
                (sanitized / "a.xlsx").write_bytes(b"dummy-xlsx")
                (originals / "b.xlsx.enc").write_bytes(b"dummy-enc")

                expectations_path = Path(tmp) / "expectations.json"
                expectations_path.write_text(json.dumps({"a.xlsx": {"open_ok": True}}), encoding="utf-8")

                out_dir = Path(tmp) / "out"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.triage",
                        "--corpus-dir",
                        str(corpus_dir),
                        "--out-dir",
                        str(out_dir),
                        "--expectations",
                        str(expectations_path),
                    ]
                    with mock.patch("sys.stdout", new=io.StringIO()):
                        rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                self.assertFalse((out_dir / "expectations-result.json").exists())
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]
            triage_mod._compare_expectations = original_compare_expectations  # type: ignore[assignment]

    def test_main_all_includes_both_sanitized_and_originals(self) -> None:
        import tools.corpus.triage as triage_mod

        observed: dict[str, object] = {}

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
                observed["paths"] = list(paths)
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
                        "--input-scope",
                        "all",
                    ]
                    with mock.patch("sys.stdout", new=io.StringIO()):
                        rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                rel_paths = {
                    p.relative_to(corpus_dir).as_posix()
                    for p in observed.get("paths", [])
                }
                self.assertEqual(rel_paths, {"sanitized/a.xlsx", "originals/b.xlsx.enc"})

                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertEqual(index["input_scope"], "all")
                self.assertEqual(index["input_dir"], str(corpus_dir))
                self.assertEqual(index["report_count"], 2)
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]

    def test_main_originals_scope_enumerates_originals_when_key_present(self) -> None:
        import tools.corpus.triage as triage_mod

        observed: dict[str, object] = {}

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
                observed["paths"] = list(paths)
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
                originals = corpus_dir / "originals"
                originals.mkdir(parents=True)
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
                        "--input-scope",
                        "originals",
                    ]
                    with mock.patch.dict("os.environ", {"CORPUS_ENCRYPTION_KEY": "key"}, clear=True):
                        with mock.patch("sys.stdout", new=io.StringIO()):
                            rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                rel_paths = {
                    p.relative_to(corpus_dir).as_posix()
                    for p in observed.get("paths", [])
                }
                self.assertEqual(rel_paths, {"originals/b.xlsx.enc"})

                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertEqual(index["input_scope"], "originals")
                self.assertEqual(index["input_dir"], str(originals))
                self.assertEqual(index["report_count"], 1)
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]

    def test_main_sanitized_missing_errors(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory() as tmp:
            corpus_dir = Path(tmp) / "corpus"
            corpus_dir.mkdir(parents=True)
            out_dir = Path(tmp) / "out"

            argv = sys.argv
            try:
                sys.argv = [
                    "tools.corpus.triage",
                    "--corpus-dir",
                    str(corpus_dir),
                    "--out-dir",
                    str(out_dir),
                    "--input-scope",
                    "sanitized",
                ]
                with mock.patch("sys.stderr", new=io.StringIO()):
                    with self.assertRaises(SystemExit) as ctx:
                        triage_mod.main()
            finally:
                sys.argv = argv

            self.assertEqual(ctx.exception.code, 2)

    def test_main_originals_requires_decryption_key(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory() as tmp:
            corpus_dir = Path(tmp) / "corpus"
            originals = corpus_dir / "originals"
            originals.mkdir(parents=True)
            (originals / "book.xlsx.enc").write_bytes(b"dummy")
            out_dir = Path(tmp) / "out"

            argv = sys.argv
            try:
                sys.argv = [
                    "tools.corpus.triage",
                    "--corpus-dir",
                    str(corpus_dir),
                    "--out-dir",
                    str(out_dir),
                    "--input-scope",
                    "originals",
                ]
                with mock.patch.dict("os.environ", {}, clear=True):
                    with mock.patch("sys.stderr", new=io.StringIO()):
                        with self.assertRaises(SystemExit) as ctx:
                            triage_mod.main()
            finally:
                sys.argv = argv

            self.assertEqual(ctx.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
