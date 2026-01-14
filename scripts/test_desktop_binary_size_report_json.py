from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class DesktopBinarySizeReportJsonTests(unittest.TestCase):
    def _repo_root(self) -> Path:
        return Path(__file__).resolve().parents[1]

    def _script_path(self) -> Path:
        script_py = Path(__file__).resolve().parent / "desktop_binary_size_report.py"
        self.assertTrue(script_py.is_file(), f"desktop_binary_size_report.py not found at {script_py}")
        return script_py

    def _run(
        self,
        repo_root: Path,
        argv: list[str],
        *,
        extra_env: dict[str, str] | None = None,
    ) -> subprocess.CompletedProcess[str]:
        env = os.environ.copy()
        env.pop("FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB", None)
        env.pop("FORMULA_ENFORCE_DESKTOP_BINARY_SIZE", None)
        env.pop("GITHUB_STEP_SUMMARY", None)
        env["RUNNER_OS"] = "UnitTestOS"
        if extra_env:
            env.update(extra_env)

        return subprocess.run(
            [sys.executable, str(self._script_path()), *argv],
            cwd=repo_root,
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )

    def _read_report(self, json_path: Path) -> dict:
        self.assertTrue(json_path.is_file(), f"Expected JSON report at {json_path}")
        return json.loads(json_path.read_text(encoding="utf-8"))

    def test_json_schema_ok_with_synthetic_binary(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            cargo_target_dir = tmp_dir / "cargo-target"
            (cargo_target_dir / "release").mkdir(parents=True, exist_ok=True)

            # Create a synthetic "desktop binary" at the expected location so the report can run
            # without performing a full Tauri build.
            fake_bin = cargo_target_dir / "release" / "formula-desktop"
            with open(fake_bin, "wb") as f:
                f.truncate(1234)

            json_path = tmp_dir / "desktop-binary-size.json"
            proc = self._run(
                repo_root,
                ["--no-build", "--json-out", str(json_path)],
                extra_env={"CARGO_TARGET_DIR": str(cargo_target_dir)},
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertIn("## Desktop Rust binary size breakdown", proc.stdout)

            report = self._read_report(json_path)
            self.assertEqual(report.get("status"), "ok")
            self.assertEqual(report.get("bin_name"), "formula-desktop")
            self.assertEqual(report.get("bin_path"), fake_bin.relative_to(repo_root).as_posix())
            self.assertEqual(report.get("bin_size_bytes"), 1234)
            self.assertIn("generated_at", report)
            self.assertIn("toolchain", report)
            self.assertIn("commands", report)

    def test_json_schema_written_when_binary_missing(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        # If a developer has already built a real desktop binary locally, the report script may
        # find it via the fallback search directories (repo-root target/, src-tauri/target). In
        # that case this test would become non-deterministic; skip it.
        fallback_candidates = [
            repo_root / "target" / "release" / "formula-desktop",
            repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "formula-desktop",
        ]
        if any(p.is_file() for p in fallback_candidates):
            self.skipTest("Desktop binary already exists in repo; skipping missing-binary JSON test.")

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            cargo_target_dir = tmp_dir / "cargo-target"
            cargo_target_dir.mkdir(parents=True, exist_ok=True)

            json_path = tmp_dir / "desktop-binary-size.json"
            proc = self._run(
                repo_root,
                ["--no-build", "--json-out", str(json_path)],
                extra_env={"CARGO_TARGET_DIR": str(cargo_target_dir)},
            )
            self.assertEqual(proc.returncode, 1)

            report = self._read_report(json_path)
            self.assertEqual(report.get("status"), "error")
            self.assertEqual(report.get("error"), "binary not found")
            self.assertIn("searched_paths", report)
            self.assertIsInstance(report["searched_paths"], list)

