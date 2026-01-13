from __future__ import annotations

import json
import subprocess
import unittest
from pathlib import Path

import tools.corpus.triage as triage_mod


class TriageDiffIgnoreGlobTests(unittest.TestCase):
    def test_run_rust_triage_passes_ignore_globs_to_helper(self) -> None:
        captured_cmd: list[str] | None = None

        original_run = triage_mod.subprocess.run
        try:

            def fake_run(cmd: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
                nonlocal captured_cmd
                captured_cmd = cmd
                return subprocess.CompletedProcess(
                    cmd,
                    0,
                    stdout=json.dumps(
                        {"steps": {}, "result": {"open_ok": True, "round_trip_ok": True}}
                    ),
                    stderr="",
                )

            triage_mod.subprocess.run = fake_run  # type: ignore[assignment]

            triage_mod._run_rust_triage(
                Path("formula-corpus-triage"),
                b"dummy",
                workbook_name="book.xlsx",
                diff_ignore={"docProps/core.xml"},
                diff_ignore_globs={"xl/media/*"},
                diff_limit=25,
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod.subprocess.run = original_run  # type: ignore[assignment]

        self.assertIsNotNone(captured_cmd)
        assert captured_cmd is not None  # for type narrowing
        # The ignore glob should be passed through to the Rust helper as a repeated flag.
        pairs = [captured_cmd[i : i + 2] for i in range(len(captured_cmd) - 1)]
        self.assertIn(["--ignore-glob", "xl/media/*"], pairs)


if __name__ == "__main__":
    unittest.main()
