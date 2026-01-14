from __future__ import annotations

import io
import tempfile
import time
import unittest
from concurrent.futures import ThreadPoolExecutor
from contextlib import redirect_stdout
from pathlib import Path

from tools.corpus.util import sha256_hex


class TriageParallelSchedulingTests(unittest.TestCase):
    def test_parallel_triage_returns_reports_in_input_order(self) -> None:
        import tools.corpus.triage as triage_mod

        original_triage_workbook = triage_mod.triage_workbook
        try:
            def fake_triage_workbook(workbook, **_kwargs):  # type: ignore[no-untyped-def]
                # Sleep based on name to force out-of-order completions.
                if workbook.display_name == "a.xlsx":
                    time.sleep(0.05)
                elif workbook.display_name == "b.xlsx":
                    time.sleep(0.02)
                else:
                    time.sleep(0.0)
                return {
                    "display_name": workbook.display_name,
                    "sha256": sha256_hex(workbook.data),
                    "result": {"open_ok": True},
                }

            triage_mod.triage_workbook = fake_triage_workbook  # type: ignore[assignment]

            with tempfile.TemporaryDirectory() as td:
                corpus_dir = Path(td) / "corpus"
                corpus_dir.mkdir(parents=True)
                (corpus_dir / "a.xlsx").write_bytes(b"a")
                (corpus_dir / "b.xlsx").write_bytes(b"b")
                (corpus_dir / "c.xlsx").write_bytes(b"c")

                paths = list(triage_mod.iter_workbook_paths(corpus_dir))
                buf = io.StringIO()
                with redirect_stdout(buf):
                    out = triage_mod._triage_paths(
                        paths,
                        rust_exe="noop",
                        diff_ignore=set(),
                        diff_limit=0,
                        recalc=False,
                        render_smoke=False,
                        leak_scan=False,
                        fernet_key=None,
                        jobs=3,
                        executor_cls=ThreadPoolExecutor,
                    )

                # _triage_paths returns an ordered list, even though completion order differed.
                self.assertIsInstance(out, list)
                reports = out
                self.assertEqual([r["display_name"] for r in reports], ["a.xlsx", "b.xlsx", "c.xlsx"])
        finally:
            triage_mod.triage_workbook = original_triage_workbook  # type: ignore[assignment]

    def test_report_filenames_are_deterministic_and_non_colliding(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory() as td:
            corpus_dir = Path(td) / "corpus"
            corpus_dir.mkdir(parents=True)
            path_a = corpus_dir / "a.xlsx"
            path_b = corpus_dir / "b.xlsx"
            path_a.write_bytes(b"same")
            path_b.write_bytes(b"same")

            report = {"sha256": sha256_hex(b"same")}

            name_a_1 = triage_mod._report_filename_for_path(report, path=path_a, corpus_dir=corpus_dir)
            name_a_2 = triage_mod._report_filename_for_path(report, path=path_a, corpus_dir=corpus_dir)
            name_b = triage_mod._report_filename_for_path(report, path=path_b, corpus_dir=corpus_dir)

            self.assertEqual(name_a_1, name_a_2)
            self.assertNotEqual(name_a_1, name_b)
            self.assertTrue(name_a_1.endswith(".json"))

    def test_jobs_produce_same_report_filenames(self) -> None:
        """`--jobs 1` and `--jobs > 1` should emit the same report file set.

        This test exercises the scheduling layer only (no Rust) by mocking `triage_workbook`
        and using a thread-based executor for the parallel branch.
        """

        import tools.corpus.triage as triage_mod

        original_triage_workbook = triage_mod.triage_workbook
        try:
            triage_mod.triage_workbook = (  # type: ignore[assignment]
                lambda wb, **_kwargs: {
                    "display_name": wb.display_name,
                    "sha256": sha256_hex(wb.data),
                    "result": {"open_ok": True, "round_trip_ok": True},
                }
            )

            with tempfile.TemporaryDirectory() as td:
                corpus_dir = Path(td) / "corpus"
                (corpus_dir / "nested").mkdir(parents=True)
                (corpus_dir / "a.xlsx").write_bytes(b"same")
                (corpus_dir / "nested" / "a.xlsx").write_bytes(b"same")
                (corpus_dir / "b.xlsx").write_bytes(b"b")

                paths = list(triage_mod.iter_workbook_paths(corpus_dir))

                with redirect_stdout(io.StringIO()):
                    out_serial = triage_mod._triage_paths(
                        paths,
                        rust_exe="noop",
                        diff_ignore=set(),
                        diff_limit=0,
                        recalc=False,
                        render_smoke=False,
                        leak_scan=False,
                        fernet_key=None,
                        jobs=1,
                    )
                    out_parallel = triage_mod._triage_paths(
                        paths,
                        rust_exe="noop",
                        diff_ignore=set(),
                        diff_limit=0,
                        recalc=False,
                        render_smoke=False,
                        leak_scan=False,
                        fernet_key=None,
                        jobs=3,
                        executor_cls=ThreadPoolExecutor,
                    )

                self.assertIsInstance(out_serial, list)
                self.assertIsInstance(out_parallel, list)

                serial_files = [
                    triage_mod._report_filename_for_path(r, path=p, corpus_dir=corpus_dir)
                    for p, r in zip(paths, out_serial, strict=True)
                ]
                parallel_files = [
                    triage_mod._report_filename_for_path(r, path=p, corpus_dir=corpus_dir)
                    for p, r in zip(paths, out_parallel, strict=True)
                ]

                self.assertEqual(serial_files, parallel_files)
                self.assertEqual(len(serial_files), len(set(serial_files)))
        finally:
            triage_mod.triage_workbook = original_triage_workbook  # type: ignore[assignment]

    def test_parallel_progress_logs_do_not_leak_raw_filenames_in_private_mode(self) -> None:
        """Defense-in-depth: even if a report is missing display_name, private mode should avoid leaks.

        Progress logs are printed by the parent process in `_triage_paths` for parallel runs. These logs
        may be captured by CI and uploaded alongside artifacts, so they must not include raw filenames
        for private corpora.
        """

        import tools.corpus.triage as triage_mod

        original_triage_workbook = triage_mod.triage_workbook
        try:
            # Return a minimal report *without* display_name so `_display_name_for_report` must use its
            # privacy-mode fallback logic.
            triage_mod.triage_workbook = (  # type: ignore[assignment]
                lambda wb, **_kwargs: {
                    "sha256": sha256_hex(wb.data),
                    "result": {"open_ok": True, "round_trip_ok": True},
                }
            )

            with tempfile.TemporaryDirectory() as td:
                corpus_dir = Path(td) / "corpus"
                corpus_dir.mkdir(parents=True)
                (corpus_dir / "sensitive-a.xlsx").write_bytes(b"a")
                (corpus_dir / "sensitive-b.xlsx").write_bytes(b"b")

                paths = list(triage_mod.iter_workbook_paths(corpus_dir))
                buf = io.StringIO()
                with redirect_stdout(buf):
                    out = triage_mod._triage_paths(
                        paths,
                        rust_exe="noop",
                        diff_ignore=set(),
                        diff_limit=0,
                        recalc=False,
                        render_smoke=False,
                        leak_scan=False,
                        fernet_key=None,
                        jobs=2,
                        privacy_mode="private",
                        executor_cls=ThreadPoolExecutor,
                    )

                self.assertIsInstance(out, list)

                stdout = buf.getvalue()
                self.assertNotIn("sensitive-a.xlsx", stdout)
                self.assertNotIn("sensitive-b.xlsx", stdout)
                self.assertIn("workbook-", stdout)
        finally:
            triage_mod.triage_workbook = original_triage_workbook  # type: ignore[assignment]


if __name__ == "__main__":
    unittest.main()
