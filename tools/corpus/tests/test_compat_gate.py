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

    def test_calc_rate_uses_attempted_denominator(self) -> None:
        # When calculate is skipped for some workbooks, the rate should be computed among
        # attempted workbooks only (not total corpus size).
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 8,
                            "calculate_attempted": 8,
                            "render_ok": 0,
                            "round_trip_ok": 10,
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
                        "--min-calc-rate",
                        "0.90",
                    ]
                )
            self.assertEqual(rc, 0, buf.getvalue())

    def test_render_rate_uses_attempted_denominator(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 0,
                            "render_ok": 7,
                            "render_attempted": 7,
                            "round_trip_ok": 10,
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
                        "--min-render-rate",
                        "0.90",
                    ]
                )
            self.assertEqual(rc, 0, buf.getvalue())

    def test_calc_render_violation_details_use_attempted_denominators(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 7,
                            "calculate_attempted": 8,
                            "render_ok": 8,
                            "render_attempted": 9,
                            "round_trip_ok": 10,
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
                        "--min-calc-rate",
                        "0.90",
                        "--min-render-rate",
                        "0.90",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 1, out)
            # Violations should be reported as ok/attempted, not ok/total.
            first_line = out.splitlines()[0] if out else ""
            self.assertIn("calculate 7/8", first_line)
            self.assertIn("render 8/9", first_line)
            self.assertNotIn("calculate 7/10", out)
            self.assertNotIn("render 8/10", out)

    def test_calc_attempted_zero_is_configuration_error_when_threshold_requested(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 0,
                            "calculate_attempted": 0,
                            "render_ok": 0,
                            "round_trip_ok": 10,
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
                        "--min-calc-rate",
                        "0.01",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 2, out)
            self.assertIn("CORPUS GATE ERROR", out)
            self.assertIn("calculate_attempted=0", out)

    def test_render_attempted_zero_is_configuration_error_when_threshold_requested(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 0,
                            "render_ok": 0,
                            "render_attempted": 0,
                            "round_trip_ok": 10,
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
                        "--min-render-rate",
                        "0.01",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 2, out)
            self.assertIn("CORPUS GATE ERROR", out)
            self.assertIn("render_attempted=0", out)

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

    def test_gate_uses_attempted_denominator_for_calc_and_render(self) -> None:
        # Calculate/render are optional steps. Gates should use the attempted denominators rather
        # than total workbooks (otherwise enabling recalc/render would look like massive regressions
        # when many workbooks legitimately skip those checks).
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 5,
                            "calculate_attempted": 5,
                            "render_ok": 2,
                            "render_attempted": 2,
                            "round_trip_ok": 10,
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
                        "--min-calc-rate",
                        "1.0",
                        "--min-render-rate",
                        "1.0",
                    ]
                )
            self.assertEqual(rc, 0, buf.getvalue())

    def test_gate_errors_when_calc_gate_is_set_but_no_attempts(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 0,
                            "calculate_attempted": 0,
                            "render_ok": 0,
                            "render_attempted": 0,
                            "round_trip_ok": 10,
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
                        "--min-calc-rate",
                        "1.0",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 2, out)
            self.assertIn("no calculate results were attempted", out)

    def test_calc_cell_fidelity_gate_passes(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 10,
                            "calculate_attempted": 10,
                            "render_ok": 0,
                            "round_trip_ok": 10,
                        },
                        "calculate_cells": {
                            "formula_cells": 1000,
                            "mismatched_cells": 1,
                            "fidelity": 0.999,
                        },
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
                        "--min-calc-cell-fidelity",
                        "0.999",
                    ]
                )
            self.assertEqual(rc, 0, buf.getvalue())

    def test_calc_cell_fidelity_gate_fails(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 10,
                            "calculate_attempted": 10,
                            "render_ok": 0,
                            "round_trip_ok": 10,
                        },
                        "calculate_cells": {
                            "formula_cells": 1000,
                            "mismatched_cells": 5,
                            "fidelity": 0.995,
                        },
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
                        "--min-calc-cell-fidelity",
                        "0.999",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 1, out)
            self.assertIn("calc-cell-fidelity", out)

    def test_calc_cell_fidelity_gate_errors_when_unavailable(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            summary_path = Path(td) / "summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "counts": {
                            "total": 10,
                            "open_ok": 10,
                            "calculate_ok": 10,
                            "calculate_attempted": 10,
                            "render_ok": 0,
                            "round_trip_ok": 10,
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
                        "--min-calc-cell-fidelity",
                        "0.999",
                    ]
                )
            out = buf.getvalue()
            self.assertEqual(rc, 2, out)
            self.assertIn("CORPUS GATE ERROR", out)


if __name__ == "__main__":
    unittest.main()
