from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


class TriageRayonTuningTests(unittest.TestCase):
    def test_main_sets_rayon_num_threads_based_on_effective_jobs(self) -> None:
        import tools.corpus.triage as triage_mod

        observed: dict[str, object] = {}

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **kwargs):  # type: ignore[no-untyped-def]
                observed["rayon"] = os.environ.get("RAYON_NUM_THREADS")
                observed["jobs_arg"] = kwargs.get("jobs")
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

            with tempfile.TemporaryDirectory(prefix="corpus-triage-rayon-") as td:
                corpus_dir = Path(td) / "corpus"
                corpus_dir.mkdir(parents=True)
                # Two workbooks; request more jobs than inputs.
                (corpus_dir / "a.xlsx").write_bytes(b"a")
                (corpus_dir / "b.xlsx").write_bytes(b"b")

                out_dir = Path(td) / "out"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.triage",
                        "--corpus-dir",
                        str(corpus_dir),
                        "--out-dir",
                        str(out_dir),
                        "--jobs",
                        "99",
                    ]
                    with mock.patch.dict(os.environ, {}, clear=True):
                        with mock.patch("os.cpu_count", return_value=8):
                            with mock.patch("sys.stdout", new=io.StringIO()):
                                rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                self.assertEqual(observed.get("jobs_arg"), 99)
                # Effective workers = min(99, 2) = 2; cpu=8 => rayon=4.
                self.assertEqual(observed.get("rayon"), "4")

                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertEqual(index.get("jobs"), 99)
                self.assertEqual(index.get("jobs_effective"), 2)
                self.assertEqual(index.get("rayon_num_threads"), 4)
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]


if __name__ == "__main__":
    unittest.main()

