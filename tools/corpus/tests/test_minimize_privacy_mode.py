from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from tools.corpus.util import sha256_hex


def _make_empty_zip() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED):
        pass
    return buf.getvalue()


class MinimizePrivacyModeTests(unittest.TestCase):
    def test_minimize_private_mode_anonymizes_display_name_and_hashes_run_url(self) -> None:
        import tools.corpus.minimize as minimize_mod
        import tools.corpus.triage as triage_mod

        original_build_rust_helper = triage_mod._build_rust_helper
        original_run_rust_triage = triage_mod._run_rust_triage
        original_env = os.environ.copy()
        argv = sys.argv
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {"diff": {"details": {"counts": {"critical": 0, "warning": 0, "info": 0, "total": 0}}}},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            os.environ["GITHUB_SERVER_URL"] = "https://github.corp.example.com"
            os.environ["GITHUB_REPOSITORY"] = "corp/repo"
            os.environ["GITHUB_RUN_ID"] = "123"

            with tempfile.TemporaryDirectory() as td:
                td_path = Path(td)
                input_path = td_path / "sensitive-name.xlsx"
                input_path.write_bytes(_make_empty_zip())
                out_path = td_path / "out.json"

                sys.argv = [
                    "tools.corpus.minimize",
                    "--input",
                    str(input_path),
                    "--out",
                    str(out_path),
                    "--privacy-mode",
                    "private",
                ]
                with mock.patch("sys.stdout", new=io.StringIO()):
                    rc = minimize_mod.main()
                self.assertEqual(rc, 0)

                out = json.loads(out_path.read_text(encoding="utf-8"))
                expected_sha = sha256_hex(input_path.read_bytes())
                self.assertEqual(out["display_name"], f"workbook-{expected_sha[:16]}.xlsx")
                self.assertTrue(out["run_url"].startswith("sha256="))
                self.assertNotIn("github.corp.example.com", out["run_url"])
        finally:
            sys.argv = argv
            os.environ.clear()
            os.environ.update(original_env)
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]


if __name__ == "__main__":
    unittest.main()
