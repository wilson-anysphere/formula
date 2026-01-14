from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class DesktopBundleSizeReportJsonTests(unittest.TestCase):
    def _script_path(self) -> Path:
        script_py = Path(__file__).resolve().parent / "desktop_bundle_size_report.py"
        self.assertTrue(script_py.is_file(), f"desktop_bundle_size_report.py not found at {script_py}")
        return script_py

    def test_src_tauri_discovery_is_depth_bounded(self) -> None:
        """
        Perf guardrail: avoid unbounded `os.walk(repo_root)` scans when falling back to
        src-tauri discovery (repo may contain large build output trees in CI).
        """
        src = self._script_path().read_text(encoding="utf-8")
        # We don't parse Python; just assert the expected max-depth guard exists.
        self.assertIn("max_depth = 8", src)
        self.assertIn("if depth >= max_depth", src)

    def _run(self, repo_root: Path, argv: list[str]) -> subprocess.CompletedProcess[str]:
        env = os.environ.copy()
        env.pop("FORMULA_BUNDLE_SIZE_JSON_PATH", None)
        env.pop("FORMULA_ENFORCE_BUNDLE_SIZE", None)
        env["RUNNER_OS"] = "UnitTestOS"

        return subprocess.run(
            [sys.executable, str(self._script_path()), *argv],
            cwd=repo_root,
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )

    def _read_report(self, repo_root: Path, rel_path: Path) -> dict:
        report_path = repo_root / rel_path
        self.assertTrue(report_path.is_file(), f"Expected JSON report at {report_path}")
        return json.loads(report_path.read_text(encoding="utf-8"))

    def _assert_basic_schema(self, report: dict) -> None:
        self.assertIsInstance(report, dict)
        self.assertIn("limit_mb", report)
        self.assertIsInstance(report["limit_mb"], int)
        self.assertIn("enforce", report)
        self.assertIsInstance(report["enforce"], bool)
        self.assertIn("bundle_dirs", report)
        self.assertIsInstance(report["bundle_dirs"], list)
        for d in report["bundle_dirs"]:
            self.assertIsInstance(d, str)
        self.assertIn("artifacts", report)
        self.assertIsInstance(report["artifacts"], list)
        self.assertIn("total_artifacts", report)
        self.assertIsInstance(report["total_artifacts"], int)
        self.assertIn("over_limit_count", report)
        self.assertIsInstance(report["over_limit_count"], int)
        # Optional key when present in the environment.
        self.assertEqual(report.get("runner_os"), "UnitTestOS")

    def test_writes_json_even_without_bundle_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            repo_root = Path(tmp_dir)
            json_rel = Path("target") / "bundle-size.json"

            proc = self._run(repo_root, ["--json", str(json_rel), "--limit-mb", "12"])
            self.assertEqual(proc.returncode, 1)
            self.assertNotIn("## Desktop installer artifact sizes", proc.stdout)

            report = self._read_report(repo_root, json_rel)
            self._assert_basic_schema(report)
            self.assertEqual(report["limit_mb"], 12)
            self.assertEqual(report["bundle_dirs"], [])
            self.assertEqual(report["total_artifacts"], 0)
            self.assertEqual(report["over_limit_count"], 0)
            self.assertEqual(report["artifacts"], [])

    def test_json_flag_defaults_to_standard_filename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            repo_root = Path(tmp_dir)
            default_path = Path("desktop-bundle-size-report.json")

            proc = self._run(repo_root, ["--json", "--limit-mb", "12"])
            self.assertEqual(proc.returncode, 1)

            report = self._read_report(repo_root, default_path)
            self._assert_basic_schema(report)
            self.assertEqual(report["limit_mb"], 12)
            self.assertEqual(report["bundle_dirs"], [])

    def test_json_schema_contains_artifacts_and_over_limit(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            repo_root = Path(tmp_dir)
            bundle_dir = repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "bundle"
            bundle_dir.mkdir(parents=True, exist_ok=True)

            # Create a synthetic artifact that exceeds a 1MB limit.
            artifact_path = bundle_dir / "formula.dmg"
            with open(artifact_path, "wb") as f:
                f.truncate(1_500_000)

            json_rel = Path("target") / "bundle-size.json"

            proc = self._run(repo_root, ["--json", str(json_rel), "--limit-mb", "1"])
            self.assertEqual(proc.returncode, 0)
            self.assertIn("## Desktop installer artifact sizes", proc.stdout)
            # Ensure the human-facing markdown table is rendered (this is also what gets
            # appended to GITHUB_STEP_SUMMARY in CI).
            self.assertIn("| Artifact | Size | Over limit |", proc.stdout)
            self.assertIn("apps/desktop/src-tauri/target/release/bundle/formula.dmg", proc.stdout)

            report = self._read_report(repo_root, json_rel)
            self._assert_basic_schema(report)
            self.assertEqual(report["limit_mb"], 1)
            self.assertFalse(report["enforce"])
            self.assertEqual(report["bundle_dirs"], ["apps/desktop/src-tauri/target/release/bundle"])

            self.assertEqual(report["total_artifacts"], 1)
            self.assertEqual(report["over_limit_count"], 1)

            art = report["artifacts"][0]
            self.assertEqual(art["path"], "apps/desktop/src-tauri/target/release/bundle/formula.dmg")
            self.assertEqual(art["size_bytes"], 1_500_000)
            self.assertTrue(art["over_limit"])
            self.assertAlmostEqual(art["size_mb"], 1.5, places=3)

    def test_enforcement_failure_still_writes_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            repo_root = Path(tmp_dir)
            bundle_dir = repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / "bundle"
            bundle_dir.mkdir(parents=True, exist_ok=True)

            artifact_path = bundle_dir / "formula.dmg"
            with open(artifact_path, "wb") as f:
                f.truncate(1_500_000)

            json_rel = Path("target") / "bundle-size.json"
            proc = self._run(repo_root, ["--json", str(json_rel), "--limit-mb", "1", "--enforce"])
            self.assertEqual(proc.returncode, 1)
            # When enforcement is enabled, the script should fail with a clear error listing the offender(s).
            self.assertIn("bundle-size: ERROR", proc.stderr)
            self.assertIn("apps/desktop/src-tauri/target/release/bundle/formula.dmg", proc.stderr)

            report = self._read_report(repo_root, json_rel)
            self._assert_basic_schema(report)
            self.assertTrue(report["enforce"])
            self.assertEqual(report["bundle_dirs"], ["apps/desktop/src-tauri/target/release/bundle"])
            self.assertEqual(report["over_limit_count"], 1)
