from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


class DashboardIndexMetadataTests(unittest.TestCase):
    def test_dashboard_uses_index_commit_and_run_url_when_env_missing(self) -> None:
        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory(prefix="corpus-dashboard-index-meta-") as td:
            triage_dir = Path(td) / "triage"
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True)

            # Minimal report set.
            (reports_dir / "r.json").write_text(
                json.dumps({"display_name": "book.xlsx", "result": {"open_ok": True, "round_trip_ok": True}}),
                encoding="utf-8",
            )

            (triage_dir / "index.json").write_text(
                json.dumps({"commit": "deadbeef", "run_url": "https://example.invalid/run/1"}),
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
                ]
                # Ensure the env-based metadata is missing so dashboard must fall back to index.json.
                with mock.patch.dict(os.environ, {}, clear=True):
                    with mock.patch("sys.stdout", new=io.StringIO()):
                        rc = dashboard_mod.main()
            finally:
                sys.argv = argv

            self.assertEqual(rc, 0)
            summary = json.loads((out_dir / "summary.json").read_text(encoding="utf-8"))
            self.assertEqual(summary.get("commit"), "deadbeef")
            self.assertEqual(summary.get("run_url"), "https://example.invalid/run/1")


if __name__ == "__main__":
    unittest.main()
