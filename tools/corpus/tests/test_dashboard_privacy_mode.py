from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


class DashboardPrivacyModeTests(unittest.TestCase):
    def test_dashboard_private_mode_hashes_non_github_run_url(self) -> None:
        import tools.corpus.dashboard as dashboard_mod

        # Minimal triage dir with a single report so dashboard has something to summarize.
        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td) / "triage"
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True)
            (reports_dir / "r.json").write_text(
                json.dumps(
                    {
                        "display_name": "book.xlsx",
                        "result": {"open_ok": True, "round_trip_ok": True},
                    }
                ),
                encoding="utf-8",
            )

            out_dir = Path(td) / "out"

            original_env = os.environ.copy()
            argv = sys.argv
            try:
                os.environ["GITHUB_SERVER_URL"] = "https://github.corp.example.com"
                os.environ["GITHUB_REPOSITORY"] = "corp/repo"
                os.environ["GITHUB_RUN_ID"] = "123"

                sys.argv = [
                    "tools.corpus.dashboard",
                    "--triage-dir",
                    str(triage_dir),
                    "--out-dir",
                    str(out_dir),
                    "--privacy-mode",
                    "private",
                ]
                with mock.patch("sys.stdout", new=io.StringIO()):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = argv
                os.environ.clear()
                os.environ.update(original_env)

            self.assertEqual(rc, 0)
            summary = json.loads((out_dir / "summary.json").read_text(encoding="utf-8"))
            self.assertIsInstance(summary.get("run_url"), str)
            self.assertTrue(summary["run_url"].startswith("sha256="))
            self.assertNotIn("github.corp.example.com", summary["run_url"])


if __name__ == "__main__":
    unittest.main()

