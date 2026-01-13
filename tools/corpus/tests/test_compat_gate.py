from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from tools.corpus.compat_gate import main


class CompatGateTests(unittest.TestCase):
    def test_gate_passes_when_rates_meet_thresholds(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 100,
                            "open_ok": 100,
                            "calculate_ok": 0,
                            "render_ok": 0,
                            "round_trip_ok": 97,
                        }
                    }
                ),
                encoding="utf-8",
            )

            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = main(
                    [
                        "--summary-json",
                        str(summary_path),
                        "--min-round-trip-rate",
                        "0.97",
                    ]
                )
            self.assertEqual(rc, 0, buf.getvalue())

    def test_gate_fails_when_rates_drop_below_thresholds(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 100,
                            "open_ok": 99,
                            "calculate_ok": 0,
                            "render_ok": 0,
                            "round_trip_ok": 96,
                        }
                    }
                ),
                encoding="utf-8",
            )

            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = main(
                    [
                        "--summary-json",
                        str(summary_path),
                        "--min-open-rate",
                        "1.0",
                        "--min-round-trip-rate",
                        "0.97",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 1, out)
            self.assertIn("CORPUS GATE FAIL", out)
            self.assertIn("open", out)
            self.assertIn("round-trip", out)


if __name__ == "__main__":
    unittest.main()
