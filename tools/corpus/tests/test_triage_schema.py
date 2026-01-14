from __future__ import annotations

import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tools.corpus.dashboard import _markdown_summary
from tools.corpus.triage import _compare_expectations, triage_workbook
from tools.corpus.util import WorkbookInput


class TriageSchemaTests(unittest.TestCase):
    def test_compare_expectations_supports_numeric_thresholds(self) -> None:
        reports = [
            {
                "display_name": "book.xlsx",
                "result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0},
            }
        ]
        expectations = {
            "book.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}
        }
        regressions, improvements = _compare_expectations(reports, expectations)
        self.assertEqual(regressions, [])
        self.assertEqual(improvements, [])

        # Numeric regressions: higher-than-expected counts should fail CI.
        reports[0]["result"]["diff_critical_count"] = 2
        regressions, _ = _compare_expectations(reports, expectations)
        self.assertTrue(any("diff_critical_count" in r for r in regressions))

        # Numeric improvements: lower-than-expected counts are surfaced as improvements.
        expectations["book.xlsx"]["diff_critical_count"] = 2
        reports[0]["result"]["diff_critical_count"] = 0
        _, improvements = _compare_expectations(reports, expectations)
        self.assertTrue(any("diff_critical_count" in r for r in improvements))

    def test_compare_expectations_treats_skips_as_regressions(self) -> None:
        reports = [{"display_name": "book.xlsx", "result": {"open_ok": None}}]
        expectations = {"book.xlsx": {"open_ok": True}}
        regressions, _ = _compare_expectations(reports, expectations)
        self.assertEqual(len(regressions), 1)

    def test_dashboard_markdown_includes_diff_and_render_columns(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "round_trip_fail_on": "critical",
            "counts": {
                "total": 1,
                "open_ok": 1,
                "calculate_ok": 0,
                "calculate_attempted": 0,
                "render_ok": 0,
                "render_attempted": 0,
                "round_trip_ok": 1,
            },
            "rates": {"open": 1.0, "calculate": None, "render": None, "round_trip": 1.0},
            "diff_totals": {"critical": 0, "warning": 1, "info": 0},
        }
        reports = [
            {
                "display_name": "book.xlsx",
                "result": {
                    "open_ok": True,
                    "calculate_ok": None,
                    "render_ok": None,
                    "round_trip_ok": True,
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                },
            }
        ]
        md = _markdown_summary(summary, reports)
        self.assertIn("Diff (C/W/I)", md)
        self.assertIn("Round-trip kind", md)
        self.assertIn("0/1/0", md)
        self.assertIn("Round-trip fail-on", md)
        # Calculate/render should not be reported as "0.0%" when triage skipped those steps.
        self.assertIn("Calculate: **0 / 0 attempted** (SKIP", md)
        self.assertIn("Render: **0 / 0 attempted** (SKIP", md)

    def test_triage_passes_round_trip_fail_on_to_rust_and_surfaces_in_dashboard(self) -> None:
        import subprocess

        def fake_run(cmd, **kwargs):  # type: ignore[no-untyped-def]
            # Ensure the Python wrapper forwards the configured threshold.
            self.assertIn("--fail-on", cmd)
            idx = cmd.index("--fail-on")
            self.assertEqual(cmd[idx + 1], "warning")

            payload = {
                "steps": {},
                "result": {
                    "open_ok": True,
                    "round_trip_ok": False,
                    "round_trip_fail_on": "warning",
                    "diff_critical_count": 0,
                    "diff_warning_count": 1,
                    "diff_info_count": 0,
                    "diff_total_count": 1,
                },
            }
            return subprocess.CompletedProcess(
                cmd, 0, stdout=json.dumps(payload), stderr=""
            )

        with mock.patch("tools.corpus.triage.subprocess.run", side_effect=fake_run):
            report = triage_workbook(
                WorkbookInput(display_name="book.xlsx", data=b"not-a-zip"),
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=25,
                round_trip_fail_on="warning",
                recalc=False,
                render_smoke=False,
            )

        self.assertEqual(report["result"]["round_trip_fail_on"], "warning")
        self.assertFalse(report["result"]["round_trip_ok"])

        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "round_trip_fail_on": "warning",
            "counts": {
                "total": 1,
                "open_ok": 1,
                "calculate_ok": 0,
                "render_ok": 0,
                "round_trip_ok": 0,
            },
            "rates": {"open": 1.0, "calculate": 0.0, "render": 0.0, "round_trip": 0.0},
        }
        md = _markdown_summary(summary, [report])
        self.assertIn("Round-trip fail-on: `warning`", md)

    def test_index_json_records_diff_ignore_globs_and_presets(self) -> None:
        import tools.corpus.triage as triage_mod

        original_build_rust_helper = triage_mod._build_rust_helper
        original_triage_paths = triage_mod._triage_paths
        try:
            triage_mod._build_rust_helper = lambda: Path("noop")  # type: ignore[assignment]

            def _fake_triage_paths(paths, **_kwargs):  # type: ignore[no-untyped-def]
                return [
                    {
                        "display_name": p.name,
                        "sha256": "0" * 64,
                        "result": {"open_ok": True, "round_trip_ok": True},
                    }
                    for p in paths
                ]

            triage_mod._triage_paths = _fake_triage_paths  # type: ignore[assignment]

            with tempfile.TemporaryDirectory(prefix="corpus-triage-index-") as td:
                corpus_dir = Path(td) / "corpus"
                corpus_dir.mkdir(parents=True)
                (corpus_dir / "a.xlsx").write_bytes(b"a")

                out_dir = Path(td) / "out"

                argv = sys.argv
                try:
                    sys.argv = [
                        "tools.corpus.triage",
                        "--corpus-dir",
                        str(corpus_dir),
                        "--out-dir",
                        str(out_dir),
                        "--diff-ignore-glob",
                        "xl/media/*",
                        "--diff-ignore-preset",
                        "excel-volatile-ids",
                    ]
                    with mock.patch("sys.stdout", new=io.StringIO()):
                        rc = triage_mod.main()
                finally:
                    sys.argv = argv

                self.assertEqual(rc, 0)
                index = json.loads((out_dir / "index.json").read_text(encoding="utf-8"))
                self.assertIn("xl/media/*", index.get("diff_ignore_globs", []))
                self.assertEqual(index.get("diff_ignore_presets"), ["excel-volatile-ids"])
        finally:
            triage_mod._build_rust_helper = original_build_rust_helper  # type: ignore[assignment]
            triage_mod._triage_paths = original_triage_paths  # type: ignore[assignment]

    def test_dashboard_markdown_includes_style_stats_when_present(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {
                "total": 2,
                "open_ok": 2,
                "calculate_ok": 2,
                "render_ok": 2,
                "round_trip_ok": 1,
            },
            "rates": {"open": 1.0, "calculate": 1.0, "render": 1.0, "round_trip": 0.5},
            "style": {
                "cellXfs": {
                    "passing": {"count": 1, "avg": 10.0, "median": 10.0},
                    "failing": {"count": 1, "avg": 100.0, "median": 100.0},
                },
                "top_failing_by_cellXfs": [{"workbook": "bad.xlsx", "cellXfs": 100}],
            },
        }
        reports = [
            {"display_name": "good.xlsx", "result": {"open_ok": True, "round_trip_ok": True}},
            {"display_name": "bad.xlsx", "result": {"open_ok": True, "round_trip_ok": False}},
        ]
        md = _markdown_summary(summary, reports)
        self.assertIn("Style complexity (cellXfs)", md)
        self.assertIn("Top failing workbooks by cellXfs", md)
        self.assertIn("bad.xlsx", md)

    def test_dashboard_markdown_includes_top_diff_fingerprints_section(self) -> None:
        summary = {
            "timestamp": "2026-01-01T00:00:00Z",
            "counts": {"total": 1, "open_ok": 1, "calculate_ok": 1, "render_ok": 1, "round_trip_ok": 0},
            "rates": {"open": 1.0, "calculate": 1.0, "render": 1.0, "round_trip": 0.0},
            "top_diff_fingerprints_in_failures": [
                {
                    "fingerprint": "37a012601a0da63445b4fbe412c6c753406e776b652013e0ca21a56a36fb634e",
                    "count": 3,
                    "part": "xl/workbook.xml.rels",
                    "kind": "attribute_changed",
                    "path": "/Relationships/Relationship[1]/@Id",
                }
            ],
        }
        md = _markdown_summary(summary, reports=[])
        self.assertIn("Top diff fingerprints in failing workbooks", md)
        # Only the prefix should be rendered in markdown.
        self.assertIn("37a012601a0da634", md)
        self.assertNotIn(
            "37a012601a0da63445b4fbe412c6c753406e776b652013e0ca21a56a36fb634e", md
        )
        self.assertIn("xl/workbook.xml.rels", md)

    def test_dashboard_rates_use_attempted_denominator(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            triage_dir = Path(tmpdir)
            reports_dir = triage_dir / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            # One workbook attempted calculation (PASS), one skipped calculation entirely.
            (reports_dir / "a.json").write_text(
                json.dumps(
                    {
                        "display_name": "attempted.xlsx",
                        "result": {
                            "open_ok": True,
                            "calculate_ok": True,
                            "render_ok": None,
                            "round_trip_ok": True,
                        },
                    }
                ),
                encoding="utf-8",
            )
            (reports_dir / "b.json").write_text(
                json.dumps(
                    {
                        "display_name": "skipped.xlsx",
                        "result": {
                            "open_ok": True,
                            "calculate_ok": None,
                            "render_ok": None,
                            "round_trip_ok": True,
                        },
                    }
                ),
                encoding="utf-8",
            )

            subprocess.run(
                [sys.executable, "-m", "tools.corpus.dashboard", "--triage-dir", str(triage_dir)],
                check=True,
            )

            summary = json.loads((triage_dir / "summary.json").read_text(encoding="utf-8"))
            self.assertEqual(summary["counts"]["calculate_attempted"], 1)
            # Pass rate should be 1/1 among attempted, not 1/2 across all workbooks.
            self.assertEqual(summary["rates"]["calculate"], 1.0)

if __name__ == "__main__":
    unittest.main()
