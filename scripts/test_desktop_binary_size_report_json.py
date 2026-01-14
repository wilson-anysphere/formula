from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class DesktopBinarySizeReportJsonTests(unittest.TestCase):
    def _write_executable(self, path: Path, content: str) -> None:
        path.write_text(content, encoding="utf-8", newline="\n")
        path.chmod(0o755)

    def _fake_rust_toolchain_env(self, tmp_dir: Path) -> dict[str, str]:
        """
        `desktop_binary_size_report.py` invokes `cargo metadata`, `cargo --version`, and
        `rustc --version` to populate the JSON report. In lightweight CI guard jobs we
        want this unit test to be deterministic and not depend on whatever Rust tooling
        happens to be installed on the runner image.

        We therefore provide tiny stub `cargo`/`rustc` executables via PATH. The report
        script prepends `$CARGO_HOME/bin` to PATH, so we install the stubs there.
        """

        cargo_home = tmp_dir / "fake-cargo-home"
        bin_dir = cargo_home / "bin"
        bin_dir.mkdir(parents=True, exist_ok=True)

        # Keep these strings stable; the unit tests only assert they are non-empty strings.
        fake_version = "1.92.0"

        if sys.platform == "win32":
            # Windows resolves commands via PATHEXT; use .cmd wrappers.
            self._write_executable(
                bin_dir / "cargo.cmd",
                f"@echo off\r\n"
                f"python -c \"import json, os, pathlib, sys; "
                f"args=sys.argv[1:]; "
                f""
                f""
                f"\\n"
                f"def out(s): sys.stdout.write(s + '\\n'); "
                f"def err(s): sys.stderr.write(s + '\\n'); "
                f""
                f"\\n"
                f"if args==['--version']: out('cargo {fake_version} (fake)'); sys.exit(0); "
                f"if args and args[0]=='metadata': "
                f"  td=os.environ.get('CARGO_TARGET_DIR'); "
                f"  cwd=pathlib.Path(os.getcwd()); "
                f"  p=pathlib.Path(td) if td else (cwd/'target'); "
                f"  p=(cwd/p).resolve() if not p.is_absolute() else p; "
                f"  out(json.dumps({{'target_directory': str(p)}})); sys.exit(0); "
                f"if args and args[0]=='bloat': err('error: no such command: `bloat`'); sys.exit(1); "
                f"err('fake cargo: unsupported args: '+repr(args)); sys.exit(1)\" %*\r\n",
            )
            self._write_executable(
                bin_dir / "rustc.cmd",
                f"@echo off\r\n"
                f"python -c \"import sys; "
                f"args=sys.argv[1:]; "
                f""
                f"\\n"
                f"if args==['--version']: print('rustc {fake_version} (fake)'); sys.exit(0); "
                f"sys.stderr.write('fake rustc: unsupported args: '+repr(args)+'\\n'); sys.exit(1)\" %*\r\n",
            )
        else:
            self._write_executable(
                bin_dir / "cargo",
                "\n".join(
                    [
                        "#!/usr/bin/env python3",
                        "from __future__ import annotations",
                        "",
                        "import json",
                        "import os",
                        "import sys",
                        "from pathlib import Path",
                        "",
                        "fake_version = " + repr(fake_version),
                        "",
                        "",
                        "def main() -> int:",
                        "    args = sys.argv[1:]",
                        "    if args == ['--version']:",
                        "        print(f'cargo {fake_version} (fake)')",
                        "        return 0",
                        "    if args and args[0] == 'metadata':",
                        "        target_dir = os.environ.get('CARGO_TARGET_DIR')",
                        "        cwd = Path.cwd()",
                        "        td = Path(target_dir) if target_dir else (cwd / 'target')",
                        "        if not td.is_absolute():",
                        "            td = (cwd / td).resolve()",
                        "        print(json.dumps({'target_directory': str(td)}))",
                        "        return 0",
                        "    if args and args[0] == 'bloat':",
                        "        # Simulate `cargo-bloat` being absent.",
                        "        sys.stderr.write('error: no such command: `bloat`\\n')",
                        "        return 1",
                        "    sys.stderr.write(f'fake cargo: unsupported args: {args!r}\\n')",
                        "    return 1",
                        "",
                        "",
                        "if __name__ == '__main__':",
                        "    raise SystemExit(main())",
                        "",
                    ]
                ),
            )
            self._write_executable(
                bin_dir / "rustc",
                "\n".join(
                    [
                        "#!/usr/bin/env python3",
                        "from __future__ import annotations",
                        "",
                        "import sys",
                        "",
                        "fake_version = " + repr(fake_version),
                        "",
                        "",
                        "def main() -> int:",
                        "    args = sys.argv[1:]",
                        "    if args == ['--version']:",
                        "        print(f'rustc {fake_version} (fake)')",
                        "        return 0",
                        "    sys.stderr.write(f'fake rustc: unsupported args: {args!r}\\n')",
                        "    return 1",
                        "",
                        "",
                        "if __name__ == '__main__':",
                        "    raise SystemExit(main())",
                        "",
                    ]
                ),
            )

        env = {
            # Ensure the report script uses our stub toolchain.
            "CARGO_HOME": str(cargo_home),
            # Keep the existing PATH so `git` (and optional tools like `file`) still resolve.
            "PATH": f"{bin_dir}{os.pathsep}{os.environ.get('PATH','')}",
        }
        return env

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
            exe = "formula-desktop.exe" if sys.platform == "win32" else "formula-desktop"
            fake_bin = cargo_target_dir / "release" / exe
            with open(fake_bin, "wb") as f:
                f.truncate(1234)

            json_path = tmp_dir / "desktop-binary-size.json"
            proc = self._run(
                repo_root,
                ["--no-build", "--json-out", str(json_path)],
                extra_env={
                    **self._fake_rust_toolchain_env(tmp_dir),
                    "CARGO_TARGET_DIR": str(cargo_target_dir),
                },
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertIn("## Desktop Rust binary size breakdown", proc.stdout)

            report = self._read_report(json_path)
            self.assertEqual(report.get("status"), "ok")
            self.assertEqual(report.get("bin_name"), "formula-desktop")
            self.assertEqual(report.get("bin_path"), fake_bin.relative_to(repo_root).as_posix())
            self.assertEqual(report.get("bin_size_bytes"), 1234)
            self.assertIn("generated_at", report)
            self.assertEqual(report.get("build_ran"), False)
            self.assertIn("runner", report)
            self.assertEqual(report["runner"]["os"], "UnitTestOS")
            self.assertIn("toolchain", report)
            self.assertIsInstance(report["toolchain"].get("rustc"), str)
            self.assertIsInstance(report["toolchain"].get("cargo"), str)
            self.assertIn("git", report)
            self.assertIsInstance(report["git"].get("sha"), str)
            self.assertIn("commands", report)

    def test_json_schema_written_when_binary_missing(self) -> None:
        repo_root = self._repo_root()
        target_dir = repo_root / "target"
        target_dir.mkdir(parents=True, exist_ok=True)

        # If a developer has already built a real desktop binary locally, the report script may
        # find it via the fallback search directories (repo-root target/, src-tauri/target). In
        # that case this test would become non-deterministic; skip it.
        exe = "formula-desktop.exe" if sys.platform == "win32" else "formula-desktop"
        fallback_candidates = [
            repo_root / "target" / "release" / exe,
            repo_root / "apps" / "desktop" / "src-tauri" / "target" / "release" / exe,
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
                extra_env={
                    **self._fake_rust_toolchain_env(tmp_dir),
                    "CARGO_TARGET_DIR": str(cargo_target_dir),
                },
            )
            self.assertEqual(proc.returncode, 1)

            report = self._read_report(json_path)
            self.assertEqual(report.get("status"), "error")
            self.assertEqual(report.get("error"), "binary not found")
            self.assertEqual(report.get("build_ran"), False)
            self.assertIn("runner", report)
            self.assertEqual(report["runner"]["os"], "UnitTestOS")
            self.assertIn("toolchain", report)
            self.assertIsInstance(report["toolchain"].get("rustc"), str)
            self.assertIsInstance(report["toolchain"].get("cargo"), str)
            self.assertIn("git", report)
            self.assertIsInstance(report["git"].get("sha"), str)
            self.assertIn("searched_paths", report)
            self.assertIsInstance(report["searched_paths"], list)
