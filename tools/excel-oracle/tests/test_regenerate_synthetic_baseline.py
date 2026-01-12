from __future__ import annotations

import importlib.util
import os
import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest import mock


class RegenerateSyntheticBaselineDryRunTests(unittest.TestCase):
    def _load_tool(self):
        tool = Path(__file__).resolve().parents[1] / "regenerate_synthetic_baseline.py"
        self.assertTrue(tool.is_file(), f"regenerate_synthetic_baseline.py not found at {tool}")

        spec = importlib.util.spec_from_file_location("excel_oracle_regenerate_synthetic_baseline", tool)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_dry_run_does_not_invoke_subprocesses(self) -> None:
        tool = self._load_tool()

        repo_root = Path(__file__).resolve().parents[3]
        old_cwd = Path.cwd()
        old_argv = sys.argv[:]
        try:
            os.chdir(repo_root)
            sys.argv = [
                str(Path(tool.__file__)),
                "--dry-run",
                "--skip-function-catalog",
                "--skip-cases",
            ]
            buf = StringIO()
            with mock.patch.object(tool.subprocess, "run") as run_mock, redirect_stdout(buf):
                rc = tool.main()
            self.assertEqual(rc, 0)
            run_mock.assert_not_called()

            out = buf.getvalue()
            self.assertIn("+", out, "expected dry-run to print at least one command")
        finally:
            sys.argv = old_argv
            os.chdir(old_cwd)


if __name__ == "__main__":
    unittest.main()

