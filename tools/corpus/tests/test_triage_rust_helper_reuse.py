from __future__ import annotations

import io
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock


class TriageRustHelperReuseTests(unittest.TestCase):
    def test_reuses_existing_helper_when_build_fails_locally(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory(prefix="corpus-triage-helper-") as td:
            root = Path(td)
            exe = root / "target" / "debug" / triage_mod._rust_exe_name()  # noqa: SLF001
            exe.parent.mkdir(parents=True, exist_ok=True)
            exe.write_bytes(b"")

            with mock.patch.object(triage_mod, "_repo_root", return_value=root):  # noqa: SLF001
                with mock.patch.dict(os.environ, {}, clear=True):
                    with mock.patch.object(triage_mod.shutil, "which", return_value=None):
                        with mock.patch.object(
                            triage_mod.subprocess,
                            "run",
                            side_effect=subprocess.CalledProcessError(101, ["cargo", "build"]),
                        ):
                            stderr = io.StringIO()
                            with mock.patch("sys.stderr", stderr):
                                built = triage_mod._build_rust_helper()  # noqa: SLF001

            self.assertEqual(built, exe)
            self.assertIn("warning: failed to build Rust triage helper", stderr.getvalue())

    def test_does_not_reuse_existing_helper_in_ci(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory(prefix="corpus-triage-helper-") as td:
            root = Path(td)
            exe = root / "target" / "debug" / triage_mod._rust_exe_name()  # noqa: SLF001
            exe.parent.mkdir(parents=True, exist_ok=True)
            exe.write_bytes(b"")

            with mock.patch.object(triage_mod, "_repo_root", return_value=root):  # noqa: SLF001
                with mock.patch.dict(os.environ, {"CI": "1"}, clear=True):
                    with mock.patch.object(triage_mod.shutil, "which", return_value=None):
                        with mock.patch.object(
                            triage_mod.subprocess,
                            "run",
                            side_effect=subprocess.CalledProcessError(101, ["cargo", "build"]),
                        ):
                            with self.assertRaises(subprocess.CalledProcessError):
                                triage_mod._build_rust_helper()  # noqa: SLF001

    def test_raises_when_no_existing_helper(self) -> None:
        import tools.corpus.triage as triage_mod

        with tempfile.TemporaryDirectory(prefix="corpus-triage-helper-") as td:
            root = Path(td)
            # No exe file created.

            with mock.patch.object(triage_mod, "_repo_root", return_value=root):  # noqa: SLF001
                with mock.patch.dict(os.environ, {}, clear=True):
                    with mock.patch.object(triage_mod.shutil, "which", return_value=None):
                        with mock.patch.object(
                            triage_mod.subprocess,
                            "run",
                            side_effect=subprocess.CalledProcessError(101, ["cargo", "build"]),
                        ):
                            with self.assertRaises(subprocess.CalledProcessError):
                                triage_mod._build_rust_helper()  # noqa: SLF001


if __name__ == "__main__":
    unittest.main()

