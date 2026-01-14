from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


def _load_excel_oracle_script(module_name: str, script_name: str):
    tool = Path(__file__).resolve().parents[1] / script_name
    if not tool.is_file():
        raise AssertionError(f"{script_name} not found at {tool}")

    spec = importlib.util.spec_from_file_location(module_name, tool)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class ExcelOracleCiTruthinessTests(unittest.TestCase):
    def test_update_pinned_dataset_tool_env_treats_ci_zero_as_non_ci(self) -> None:
        update = _load_excel_oracle_script(
            "excel_oracle_update_pinned_dataset_ci_truthiness", "update_pinned_dataset.py"
        )

        with tempfile.TemporaryDirectory(prefix="excel-oracle-ci-truthiness-") as td:
            root = Path(td)
            home = root / "home"
            repo_root = root / "repo"
            home.mkdir(parents=True, exist_ok=True)
            repo_root.mkdir(parents=True, exist_ok=True)

            default_cargo_home = home / ".cargo"

            with mock.patch.dict(
                os.environ,
                {
                    # Ensure `Path.home()` (used to compute `~/.cargo`) resolves inside the temp dir.
                    "HOME": str(home),
                    "USERPROFILE": str(home),
                    "CI": "0",
                    "CARGO_HOME": str(default_cargo_home),
                },
                clear=True,
            ):
                env = update._tool_env(repo_root)  # noqa: SLF001 (script internal)

        self.assertEqual(Path(env["CARGO_HOME"]), repo_root / "target" / "cargo-home")

    def test_update_pinned_dataset_tool_env_treats_ci_one_as_ci(self) -> None:
        update = _load_excel_oracle_script(
            "excel_oracle_update_pinned_dataset_ci_truthiness_one", "update_pinned_dataset.py"
        )

        with tempfile.TemporaryDirectory(prefix="excel-oracle-ci-truthiness-") as td:
            root = Path(td)
            home = root / "home"
            repo_root = root / "repo"
            home.mkdir(parents=True, exist_ok=True)
            repo_root.mkdir(parents=True, exist_ok=True)

            default_cargo_home = home / ".cargo"

            with mock.patch.dict(
                os.environ,
                {
                    "HOME": str(home),
                    "USERPROFILE": str(home),
                    "CI": "1",
                    "CARGO_HOME": str(default_cargo_home),
                },
                clear=True,
            ):
                env = update._tool_env(repo_root)  # noqa: SLF001 (script internal)

        self.assertEqual(Path(env["CARGO_HOME"]), default_cargo_home)

    def test_regenerate_synthetic_baseline_tool_env_treats_ci_zero_as_non_ci(self) -> None:
        regen = _load_excel_oracle_script(
            "excel_oracle_regenerate_synthetic_baseline_ci_truthiness",
            "regenerate_synthetic_baseline.py",
        )

        with tempfile.TemporaryDirectory(prefix="excel-oracle-ci-truthiness-") as td:
            root = Path(td)
            home = root / "home"
            repo_root = root / "repo"
            home.mkdir(parents=True, exist_ok=True)
            repo_root.mkdir(parents=True, exist_ok=True)

            default_cargo_home = home / ".cargo"

            with mock.patch.dict(
                os.environ,
                {
                    "HOME": str(home),
                    "USERPROFILE": str(home),
                    "CI": "0",
                    "CARGO_HOME": str(default_cargo_home),
                },
                clear=True,
            ):
                env = regen._tool_env(repo_root)  # noqa: SLF001 (script internal)

        self.assertEqual(Path(env["CARGO_HOME"]), repo_root / "target" / "cargo-home")


if __name__ == "__main__":
    unittest.main()

