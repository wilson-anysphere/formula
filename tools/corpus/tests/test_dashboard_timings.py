from __future__ import annotations

import unittest

from tools.corpus.dashboard import _compute_timings


class DashboardTimingTests(unittest.TestCase):
    def test_compute_timings_median_and_p90(self) -> None:
        # Ten samples so the p90 is not equal to the max (interpolated percentile).
        load_durations = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]
        reports = [
            {
                "display_name": f"book-{d}.xlsx",
                "steps": {
                    "load": {"status": "ok", "duration_ms": d},
                    "round_trip": {"status": "ok", "duration_ms": d + 1},
                },
            }
            for d in load_durations
        ]
        # Failed steps should not contribute to timing percentiles.
        reports.append(
            {
                "display_name": "failed.xlsx",
                "steps": {"load": {"status": "failed", "duration_ms": 5000}},
            }
        )
        # Include a report without steps to ensure it's ignored.
        reports.append({"display_name": "nosteps.xlsx"})

        timings = _compute_timings(reports)

        load = timings["load"]
        self.assertEqual(load["count"], 10)
        self.assertEqual(load["mean_ms"], 550.0)
        self.assertEqual(load["p50_ms"], 550.0)
        self.assertEqual(load["p90_ms"], 910.0)
        self.assertEqual(load["max_ms"], 1000)

        round_trip = timings["round_trip"]
        self.assertEqual(round_trip["count"], 10)
        self.assertEqual(round_trip["p50_ms"], 551.0)
        self.assertEqual(round_trip["p90_ms"], 911.0)

        # Optional steps (recalc/render) are skipped by default and should have no samples.
        self.assertEqual(timings["recalc"]["count"], 0)
        self.assertIsNone(timings["recalc"]["p50_ms"])
        self.assertEqual(timings["render"]["count"], 0)
        self.assertIsNone(timings["render"]["p90_ms"])

    def test_gate_errors_when_no_successful_samples(self) -> None:
        import json
        import io
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            # Only failed load step -> no successful timing samples.
            (reports_dir / "r1.json").write_text(
                json.dumps({"steps": {"load": {"status": "failed", "duration_ms": 123}}}),
                encoding="utf-8",
            )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-load-p90-ms",
                    "1000",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv
            self.assertEqual(rc, 2)

    def test_gate_errors_when_no_successful_round_trip_samples(self) -> None:
        import io
        import json
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            (reports_dir / "r1.json").write_text(
                json.dumps(
                    {"steps": {"round_trip": {"status": "failed", "duration_ms": 123}}}
                ),
                encoding="utf-8",
            )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-round-trip-p90-ms",
                    "1000",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv
            self.assertEqual(rc, 2)

    def test_gate_fails_when_p90_exceeds_threshold(self) -> None:
        import io
        import json
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            # Ten samples -> load p90 = 910 (interpolated).
            for i, d in enumerate([100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]):
                (reports_dir / f"r{i}.json").write_text(
                    json.dumps({"steps": {"load": {"status": "ok", "duration_ms": d}}}),
                    encoding="utf-8",
                )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-load-p90-ms",
                    "905",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv

            self.assertEqual(rc, 1)

    def test_gate_passes_when_p90_below_threshold(self) -> None:
        import io
        import json
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            for i, d in enumerate([100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]):
                (reports_dir / f"r{i}.json").write_text(
                    json.dumps({"steps": {"load": {"status": "ok", "duration_ms": d}}}),
                    encoding="utf-8",
                )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-load-p90-ms",
                    "920",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv

            self.assertEqual(rc, 0)

    def test_round_trip_gate_fails_when_p90_exceeds_threshold(self) -> None:
        import io
        import json
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            # Ten samples -> round_trip p90 = 910 (interpolated).
            for i, d in enumerate([100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]):
                (reports_dir / f"r{i}.json").write_text(
                    json.dumps({"steps": {"round_trip": {"status": "ok", "duration_ms": d}}}),
                    encoding="utf-8",
                )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-round-trip-p90-ms",
                    "905",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv

            self.assertEqual(rc, 1)

    def test_round_trip_gate_passes_when_p90_below_threshold(self) -> None:
        import io
        import json
        import sys
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path

        import tools.corpus.dashboard as dashboard_mod

        with tempfile.TemporaryDirectory() as td:
            triage_dir = Path(td)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            for i, d in enumerate([100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]):
                (reports_dir / f"r{i}.json").write_text(
                    json.dumps({"steps": {"round_trip": {"status": "ok", "duration_ms": d}}}),
                    encoding="utf-8",
                )

            original_argv = sys.argv
            try:
                sys.argv = [
                    "dashboard.py",
                    "--triage-dir",
                    str(triage_dir),
                    "--gate-round-trip-p90-ms",
                    "920",
                ]
                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = dashboard_mod.main()
            finally:
                sys.argv = original_argv

            self.assertEqual(rc, 0)


if __name__ == "__main__":
    unittest.main()
