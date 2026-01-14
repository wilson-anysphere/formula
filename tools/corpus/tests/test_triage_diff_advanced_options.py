from __future__ import annotations

import subprocess
import unittest
from pathlib import Path
from unittest import mock

import tools.corpus.triage as triage_mod


class TriageDiffAdvancedOptionsPassthroughTests(unittest.TestCase):
    def test_run_rust_triage_passes_ignore_path_flags_and_strict_calc_chain(self) -> None:
        observed: dict[str, object] = {}

        def fake_run(cmd, **_kwargs):  # type: ignore[no-untyped-def]
            observed["cmd"] = cmd
            return subprocess.CompletedProcess(
                cmd,
                0,
                stdout='{"steps": {}, "result": {}}',
                stderr="",
            )

        with mock.patch.object(triage_mod.subprocess, "run", side_effect=fake_run):
            triage_mod._run_rust_triage(  # noqa: SLF001 (unit test)
                Path("noop"),
                b"dummy",
                workbook_name="book.xlsx",
                diff_ignore={"docProps/core.xml"},
                diff_ignore_path=("dyDescent", "xr:uid"),
                diff_ignore_path_in=("xl/worksheets/*.xml:xr:uid",),
                diff_ignore_path_kind=("attribute_changed:@",),
                diff_ignore_path_kind_in=("xl/worksheets/*.xml:attribute_changed:@",),
                diff_ignore_presets=("excel-volatile-ids",),
                diff_limit=1,
                recalc=False,
                render_smoke=False,
                strict_calc_chain=True,
            )

        cmd = observed.get("cmd")
        self.assertIsInstance(cmd, list)
        cmd_list = cmd if isinstance(cmd, list) else []

        # Unscoped ignore-path rules.
        self.assertIn("--ignore-path", cmd_list)
        self.assertIn("dyDescent", cmd_list)
        self.assertIn("xr:uid", cmd_list)

        # Scoped ignore-path-in rules.
        self.assertIn("--ignore-path-in", cmd_list)
        self.assertIn("xl/worksheets/*.xml:xr:uid", cmd_list)

        # Kind-filtered ignore-path rules.
        self.assertIn("--ignore-path-kind", cmd_list)
        self.assertIn("attribute_changed:@", cmd_list)

        # Kind-filtered scoped ignore-path-kind-in rules.
        self.assertIn("--ignore-path-kind-in", cmd_list)
        self.assertIn("xl/worksheets/*.xml:attribute_changed:@", cmd_list)

        # Strict calcChain policy.
        self.assertIn("--strict-calc-chain", cmd_list)

        # Ignore preset pass-through.
        self.assertIn("--ignore-preset", cmd_list)
        self.assertIn("excel-volatile-ids", cmd_list)


if __name__ == "__main__":
    unittest.main()
