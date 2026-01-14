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

    def test_dashboard_private_mode_anonymizes_workbook_names_and_redacts_custom_functions(self) -> None:
        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td) / "triage"
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True)

            sha = "a" * 64
            (reports_dir / "r.json").write_text(
                json.dumps(
                    {
                        "display_name": "sensitive-filename.xlsx",
                        "sha256": sha,
                        "functions": {"SUM": 1, "CORP.ADDIN.FOO": 1},
                        "result": {"open_ok": True, "round_trip_ok": False},
                        "steps": {
                            "diff": {
                                "status": "ok",
                                "details": {
                                    "top_differences": [
                                        {
                                            "fingerprint": "0" * 64,
                                            "severity": "CRITICAL",
                                            "part": "xl/workbook.xml.rels",
                                            "path": "/root@{http://corp.example.com/ns}attr",
                                            "kind": "attribute_changed",
                                        }
                                    ]
                                },
                            }
                        },
                    }
                ),
                encoding="utf-8",
            )

            out_dir = Path(td) / "out"

            argv = sys.argv
            try:
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

            self.assertEqual(rc, 0)

            md = (out_dir / "summary.md").read_text(encoding="utf-8")
            self.assertIn("workbook-aaaaaaaaaaaaaaaa.xlsx", md)
            self.assertNotIn("sensitive-filename.xlsx", md)

            summary = json.loads((out_dir / "summary.json").read_text(encoding="utf-8"))
            fns = {row["function"] for row in summary.get("top_functions_in_failures", [])}
            self.assertIn("SUM", fns)
            self.assertTrue(any(v.startswith("sha256=") for v in fns))
            self.assertNotIn("CORP.ADDIN.FOO", fns)


if __name__ == "__main__":
    unittest.main()
