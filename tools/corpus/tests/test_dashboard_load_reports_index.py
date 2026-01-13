from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.corpus.dashboard import _load_reports


class DashboardLoadReportsIndexTests(unittest.TestCase):
    def test_load_reports_prefers_index_order_when_present(self) -> None:
        with tempfile.TemporaryDirectory(prefix="corpus-dashboard-index-") as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            # Write reports whose filenames sort opposite to the order we want.
            (reports_dir / "b.json").write_text(
                json.dumps({"display_name": "B"}, indent=2, sort_keys=True),
                encoding="utf-8",
            )
            (reports_dir / "a.json").write_text(
                json.dumps({"display_name": "A"}, indent=2, sort_keys=True),
                encoding="utf-8",
            )

            # Index specifies explicit order: A then B.
            (triage_dir / "index.json").write_text(
                json.dumps(
                    {
                        "reports": [
                            {"id": "a", "display_name": "A", "file": "a.json"},
                            {"id": "b", "display_name": "B", "file": "b.json"},
                        ]
                    },
                    indent=2,
                    sort_keys=True,
                ),
                encoding="utf-8",
            )

            reports = _load_reports(reports_dir)
            self.assertEqual([r.get("display_name") for r in reports], ["A", "B"])


if __name__ == "__main__":
    unittest.main()

