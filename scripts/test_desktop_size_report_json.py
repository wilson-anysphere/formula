from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class DesktopSizeReportJsonTests(unittest.TestCase):
    def _repo_root(self) -> Path:
        return Path(__file__).resolve().parents[1]

    def _script_path(self) -> Path:
        script_py = Path(__file__).resolve().parent / "desktop_size_report.py"
        self.assertTrue(script_py.is_file(), f"desktop_size_report.py not found at {script_py}")
        return script_py

    def _run(self, repo_root: Path, argv: list[str], *, extra_env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
        env = os.environ.copy()
        env.pop("FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB", None)
        env.pop("FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB", None)
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

    def _assert_basic_schema(self, report: dict) -> None:
        self.assertIsInstance(report, dict)
        self.assertEqual(report.get("runner_os"), "UnitTestOS")

        self.assertIn("binary", report)
        self.assertIsInstance(report["binary"], dict)
        self.assertIn("dist", report)
        self.assertIsInstance(report["dist"], dict)
        self.assertIn("dist_tar_gz", report)
        self.assertIn("limits_mb", report)
        self.assertIsInstance(report["limits_mb"], dict)

        for key in ("binary", "dist"):
            entry = report[key]
            self.assertIn("path", entry)
            self.assertIsInstance(entry["path"], str)
            self.assertIn("size_bytes", entry)
            self.assertIsInstance(entry["size_bytes"], int)
            self.assertIn("size_mb", entry)
            self.assertIsInstance(entry["size_mb"], float)
            self.assertIn("over_limit", entry)
            self.assertIsInstance(entry["over_limit"], bool)

    def test_json_schema_without_gzip(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            binary_path = tmp_dir / "formula-desktop"
            dist_dir = tmp_dir / "dist"
            dist_dir.mkdir(parents=True, exist_ok=True)

            with open(binary_path, "wb") as f:
                f.truncate(1_500_000)
            (dist_dir / "a.txt").write_bytes(b"hello\n")

            json_path = tmp_dir / "desktop-size.json"
            proc = self._run(
                repo_root,
                [
                    "--binary",
                    binary_path.relative_to(repo_root).as_posix(),
                    "--dist",
                    dist_dir.relative_to(repo_root).as_posix(),
                    "--no-gzip",
                    "--json-out",
                    str(json_path),
                ],
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertIn("## Desktop size report", proc.stdout)

            report = self._read_report(json_path)
            self._assert_basic_schema(report)
            self.assertIsNone(report["dist_tar_gz"])

            self.assertEqual(
                report["binary"]["path"],
                binary_path.relative_to(repo_root).as_posix(),
            )
            self.assertEqual(report["binary"]["size_bytes"], 1_500_000)
            self.assertAlmostEqual(report["binary"]["size_mb"], 1.5, places=3)
            self.assertFalse(report["binary"]["over_limit"])

    def test_json_schema_includes_gzip_size(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            binary_path = tmp_dir / "formula-desktop"
            dist_dir = tmp_dir / "dist"
            dist_dir.mkdir(parents=True, exist_ok=True)

            with open(binary_path, "wb") as f:
                f.truncate(10)
            (dist_dir / "a.txt").write_bytes(b"hello\n")

            json_path = tmp_dir / "desktop-size.json"
            proc = self._run(
                repo_root,
                [
                    "--binary",
                    binary_path.relative_to(repo_root).as_posix(),
                    "--dist",
                    dist_dir.relative_to(repo_root).as_posix(),
                    "--json-out",
                    str(json_path),
                ],
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)

            report = self._read_report(json_path)
            self._assert_basic_schema(report)

            dist_gz = report["dist_tar_gz"]
            self.assertIsInstance(dist_gz, dict)
            self.assertGreater(dist_gz["size_bytes"], 0)
            self.assertIsInstance(dist_gz["size_mb"], float)
            self.assertFalse(report["dist"]["over_limit"])

    def test_oversize_failure_still_writes_json(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            binary_path = tmp_dir / "formula-desktop"
            dist_dir = tmp_dir / "dist"
            dist_dir.mkdir(parents=True, exist_ok=True)

            with open(binary_path, "wb") as f:
                f.truncate(1_500_000)
            (dist_dir / "a.txt").write_bytes(b"hello\n")

            json_path = tmp_dir / "desktop-size.json"
            proc = self._run(
                repo_root,
                [
                    "--binary",
                    binary_path.relative_to(repo_root).as_posix(),
                    "--dist",
                    dist_dir.relative_to(repo_root).as_posix(),
                    "--no-gzip",
                    "--json-out",
                    str(json_path),
                ],
                extra_env={"FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB": "1"},
            )
            self.assertEqual(proc.returncode, 1)

            report = self._read_report(json_path)
            self._assert_basic_schema(report)
            self.assertEqual(report["limits_mb"]["binary"], 1.0)
            self.assertEqual(report["binary"]["size_bytes"], 1_500_000)
            self.assertTrue(report["binary"]["over_limit"])

    def test_dist_oversize_failure_still_writes_json(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            binary_path = tmp_dir / "formula-desktop"
            dist_dir = tmp_dir / "dist"
            dist_dir.mkdir(parents=True, exist_ok=True)

            with open(binary_path, "wb") as f:
                f.truncate(10)
            # Create a 2MB file under dist to exceed a 1MB limit.
            with open(dist_dir / "big.bin", "wb") as f:
                f.truncate(2_000_000)

            json_path = tmp_dir / "desktop-size.json"
            proc = self._run(
                repo_root,
                [
                    "--binary",
                    binary_path.relative_to(repo_root).as_posix(),
                    "--dist",
                    dist_dir.relative_to(repo_root).as_posix(),
                    "--no-gzip",
                    "--json-out",
                    str(json_path),
                ],
                extra_env={"FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB": "1"},
            )
            self.assertEqual(proc.returncode, 1)

            report = self._read_report(json_path)
            self._assert_basic_schema(report)
            self.assertEqual(report["limits_mb"]["dist"], 1.0)
            self.assertEqual(report["dist"]["size_bytes"], 2_000_000)
            self.assertTrue(report["dist"]["over_limit"])

    def test_default_binary_auto_detects_cargo_target_dir(self) -> None:
        """
        CI runs the size report without `--binary`. Ensure it can find the binary
        in a custom CARGO_TARGET_DIR when `target/release/formula-desktop` is absent.
        """
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        exe = "formula-desktop.exe" if sys.platform == "win32" else "formula-desktop"
        default_bin = repo_root / "target" / "release" / exe
        default_bin_backup = default_bin.with_name(default_bin.name + ".bak")
        tauri_bin = repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / exe
        tauri_bin_backup = tauri_bin.with_name(tauri_bin.name + ".bak")

        moved: list[tuple[Path, Path]] = []
        try:
            # Ensure any existing binaries don't short-circuit the probe logic.
            for src, dst in ((default_bin, default_bin_backup), (tauri_bin, tauri_bin_backup)):
                if src.is_file():
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    if dst.exists():
                        dst.unlink()
                    src.rename(dst)
                    moved.append((src, dst))

            with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
                tmp_dir = Path(tmp)
                cargo_target = tmp_dir / "custom-target"
                release_dir = cargo_target / "release"
                release_dir.mkdir(parents=True, exist_ok=True)
                bin_path = release_dir / exe

                with open(bin_path, "wb") as f:
                    f.truncate(1234)

                dist_dir = tmp_dir / "dist"
                dist_dir.mkdir(parents=True, exist_ok=True)
                (dist_dir / "a.txt").write_bytes(b"hello\n")

                json_path = tmp_dir / "desktop-size.json"

                proc = self._run(
                    repo_root,
                    [
                        "--dist",
                        dist_dir.relative_to(repo_root).as_posix(),
                        "--no-gzip",
                        "--json-out",
                        str(json_path),
                    ],
                    extra_env={
                        # Intentionally provide a relative path to exercise repo-root resolution.
                        "CARGO_TARGET_DIR": cargo_target.relative_to(repo_root).as_posix(),
                    },
                )
                self.assertEqual(proc.returncode, 0, proc.stderr)

                report = self._read_report(json_path)
                self._assert_basic_schema(report)
                self.assertEqual(report["binary"]["path"], bin_path.relative_to(repo_root).as_posix())
                self.assertEqual(report["binary"]["size_bytes"], 1234)
        finally:
            for src, dst in moved:
                if src.exists():
                    # Shouldn't happen, but avoid clobbering.
                    continue
                if dst.exists():
                    dst.rename(src)

    def test_appends_markdown_to_github_step_summary(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(dir=target_dir) as tmp:
            tmp_dir = Path(tmp)
            binary_path = tmp_dir / "formula-desktop"
            dist_dir = tmp_dir / "dist"
            dist_dir.mkdir(parents=True, exist_ok=True)

            with open(binary_path, "wb") as f:
                f.truncate(10)
            (dist_dir / "a.txt").write_bytes(b"hello\n")

            json_path = tmp_dir / "desktop-size.json"
            summary_path = tmp_dir / "step-summary.md"

            proc = self._run(
                repo_root,
                [
                    "--binary",
                    binary_path.relative_to(repo_root).as_posix(),
                    "--dist",
                    dist_dir.relative_to(repo_root).as_posix(),
                    "--no-gzip",
                    "--json-out",
                    str(json_path),
                ],
                extra_env={"GITHUB_STEP_SUMMARY": str(summary_path)},
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertTrue(summary_path.is_file())
            summary = summary_path.read_text(encoding="utf-8")
            self.assertIn("## Desktop size report", summary)
            self.assertIn(binary_path.relative_to(repo_root).as_posix(), summary)
