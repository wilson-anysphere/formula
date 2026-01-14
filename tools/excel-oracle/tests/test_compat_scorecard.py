from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class CompatScorecardTests(unittest.TestCase):
    def test_merges_corpus_and_oracle_metrics(self) -> None:
        scorecard_py = Path(__file__).resolve().parents[2] / "compat_scorecard.py"
        self.assertTrue(scorecard_py.is_file(), f"compat_scorecard.py not found at {scorecard_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            corpus_path = tmp_path / "corpus-summary.json"
            oracle_path = tmp_path / "mismatch-report.json"
            out_md = tmp_path / "scorecard.md"
            out_json = tmp_path / "scorecard.json"

            corpus_payload = {
                "timestamp": "unit-test",
                "counts": {
                    "total": 10,
                    "open_ok": 10,
                    "calculate_ok": 10,
                    "render_ok": 10,
                    "round_trip_ok": 9,
                },
                "rates": {"open": 1.0, "calculate": 1.0, "render": 1.0, "round_trip": 0.9},
            }
            corpus_path.write_text(
                json.dumps(corpus_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            oracle_payload = {
                "schemaVersion": 1,
                "summary": {
                    "totalCases": 1000,
                    "mismatches": 1,
                    "mismatchRate": 0.001,
                    "maxMismatchRate": 0.01,
                    "includeTags": ["add", "sub"],
                    "excludeTags": [],
                    "maxCases": 0,
                },
            }
            oracle_path.write_text(
                json.dumps(oracle_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(scorecard_py),
                    "--corpus-summary",
                    str(corpus_path),
                    "--oracle-report",
                    str(oracle_path),
                    "--out-md",
                    str(out_md),
                    "--out-json",
                    str(out_json),
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compat_scorecard.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            md = out_md.read_text(encoding="utf-8")
            self.assertIn("includeTags: add, sub", md)
            self.assertIn("excludeTags: <none>", md)
            self.assertIn("maxCases: all", md)
            self.assertIn("| L1 | Read (corpus open) | PASS | 100.00% | 10 / 10 |", md)
            # 1 mismatch out of 1000 => 99.9% pass rate.
            self.assertIn("| L2 | Calculate (Excel oracle) | PASS | 99.90% | 999 / 1000 |", md)
            # Round-trip is 90% and target is 97% => FAIL.
            self.assertIn("| L4 | Round-trip (corpus) | FAIL | 90.00% | 9 / 10 |", md)

            payload = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertEqual(payload.get("schemaVersion"), 1)
            self.assertEqual(payload["metrics"]["l1Read"]["status"], "PASS")
            self.assertEqual(payload["metrics"]["l2Calculate"]["status"], "PASS")
            self.assertEqual(payload["metrics"]["l4RoundTrip"]["status"], "FAIL")
            self.assertAlmostEqual(payload["metrics"]["l2Calculate"]["passRate"], 0.999)
            self.assertAlmostEqual(payload["metrics"]["l2Calculate"]["mismatchRate"], 0.001)
            self.assertAlmostEqual(payload["metrics"]["l2Calculate"]["maxMismatchRate"], 0.01)
            self.assertEqual(payload["metrics"]["l2Calculate"]["passes"], 999)
            self.assertEqual(payload["metrics"]["l2Calculate"]["mismatches"], 1)
            self.assertEqual(payload["inputs"]["oracle"]["includeTags"], ["add", "sub"])
            self.assertEqual(payload["inputs"]["oracle"]["excludeTags"], [])
            self.assertEqual(payload["inputs"]["oracle"]["maxCases"], 0)

    def test_falls_back_to_counts_when_rates_are_missing(self) -> None:
        scorecard_py = Path(__file__).resolve().parents[2] / "compat_scorecard.py"
        self.assertTrue(scorecard_py.is_file(), f"compat_scorecard.py not found at {scorecard_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            corpus_path = tmp_path / "corpus-summary.json"
            oracle_path = tmp_path / "mismatch-report.json"
            out_md = tmp_path / "scorecard.md"

            # Deliberately omit `rates` to ensure the generator recomputes rates from counts.
            corpus_payload = {
                "timestamp": "unit-test",
                "counts": {
                    "total": 10,
                    "open_ok": 8,
                    "calculate_ok": 10,
                    "render_ok": 10,
                    "round_trip_ok": 7,
                },
            }
            corpus_path.write_text(
                json.dumps(corpus_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            # Deliberately omit `mismatchRate` to ensure the generator recomputes it from
            # mismatches/totalCases.
            oracle_payload = {
                "schemaVersion": 1,
                "summary": {
                    "totalCases": 100,
                    "mismatches": 5,
                },
            }
            oracle_path.write_text(
                json.dumps(oracle_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(scorecard_py),
                    "--corpus-summary",
                    str(corpus_path),
                    "--oracle-report",
                    str(oracle_path),
                    "--out-md",
                    str(out_md),
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compat_scorecard.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            md = out_md.read_text(encoding="utf-8")
            self.assertIn("| L1 | Read (corpus open) | FAIL | 80.00% | 8 / 10 |", md)
            self.assertIn("| L2 | Calculate (Excel oracle) | FAIL | 95.00% | 95 / 100 |", md)
            self.assertIn("| L4 | Round-trip (corpus) | FAIL | 70.00% | 7 / 10 |", md)

    def test_zero_totals_render_as_missing(self) -> None:
        scorecard_py = Path(__file__).resolve().parents[2] / "compat_scorecard.py"
        self.assertTrue(scorecard_py.is_file(), f"compat_scorecard.py not found at {scorecard_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            corpus_path = tmp_path / "corpus-summary.json"
            oracle_path = tmp_path / "mismatch-report.json"
            out_md = tmp_path / "scorecard.md"

            corpus_payload = {
                "timestamp": "unit-test",
                "counts": {
                    "total": 0,
                    "open_ok": 0,
                    "calculate_ok": 0,
                    "render_ok": 0,
                    "round_trip_ok": 0,
                },
            }
            corpus_path.write_text(
                json.dumps(corpus_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            oracle_payload = {
                "schemaVersion": 1,
                "summary": {
                    "totalCases": 0,
                    "mismatches": 0,
                },
            }
            oracle_path.write_text(
                json.dumps(oracle_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            out_json = tmp_path / "scorecard.json"
            proc = subprocess.run(
                [
                    sys.executable,
                    str(scorecard_py),
                    "--corpus-summary",
                    str(corpus_path),
                    "--oracle-report",
                    str(oracle_path),
                    "--out-md",
                    str(out_md),
                    "--out-json",
                    str(out_json),
                ],
                capture_output=True,
                text=True,
            )

            if proc.returncode != 0:
                self.fail(
                    f"compat_scorecard.py exited {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
                )

            md = out_md.read_text(encoding="utf-8")
            self.assertIn("| L1 | Read (corpus open) | MISSING |", md)
            self.assertIn("| L2 | Calculate (Excel oracle) | MISSING |", md)
            self.assertIn("| L4 | Round-trip (corpus) | MISSING |", md)
            self.assertIn("no cases", md)
            self.assertIn("no workbooks", md)

            payload = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertIsNone(payload["metrics"]["l2Calculate"]["mismatchRate"])
            self.assertIsNone(payload["metrics"]["l2Calculate"]["passRate"])
            self.assertEqual(payload["metrics"]["l1Read"]["status"], "MISSING")
            self.assertEqual(payload["metrics"]["l2Calculate"]["status"], "MISSING")
            self.assertEqual(payload["metrics"]["l4RoundTrip"]["status"], "MISSING")

    def test_missing_inputs_exits_nonzero(self) -> None:
        scorecard_py = Path(__file__).resolve().parents[2] / "compat_scorecard.py"
        self.assertTrue(scorecard_py.is_file(), f"compat_scorecard.py not found at {scorecard_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            corpus_path = tmp_path / "missing-corpus.json"
            oracle_path = tmp_path / "missing-oracle.json"
            out_md = tmp_path / "scorecard.md"

            proc = subprocess.run(
                [
                    sys.executable,
                    str(scorecard_py),
                    "--corpus-summary",
                    str(corpus_path),
                    "--oracle-report",
                    str(oracle_path),
                    "--out-md",
                    str(out_md),
                ],
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(proc.returncode, 0)
            self.assertIn("Missing corpus summary.json", proc.stderr)
            self.assertIn("Missing Excel-oracle mismatch report", proc.stderr)

    def test_allow_missing_inputs_renders_partial_scorecard(self) -> None:
        scorecard_py = Path(__file__).resolve().parents[2] / "compat_scorecard.py"
        self.assertTrue(scorecard_py.is_file(), f"compat_scorecard.py not found at {scorecard_py}")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            corpus_path = tmp_path / "corpus-summary.json"
            oracle_path = tmp_path / "missing-oracle.json"
            out_md = tmp_path / "scorecard.md"

            corpus_payload = {
                "timestamp": "unit-test",
                "counts": {
                    "total": 2,
                    "open_ok": 2,
                    "calculate_ok": 2,
                    "render_ok": 2,
                    "round_trip_ok": 2,
                },
                "rates": {"open": 1.0, "calculate": 1.0, "render": 1.0, "round_trip": 1.0},
            }
            corpus_path.write_text(
                json.dumps(corpus_payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
                newline="\n",
            )

            proc = subprocess.run(
                [
                    sys.executable,
                    str(scorecard_py),
                    "--corpus-summary",
                    str(corpus_path),
                    "--oracle-report",
                    str(oracle_path),
                    "--out-md",
                    str(out_md),
                    "--allow-missing-inputs",
                ],
                capture_output=True,
                text=True,
            )
            self.assertEqual(proc.returncode, 0, f"stderr:\n{proc.stderr}")
            self.assertIn("Missing Excel-oracle mismatch report", proc.stderr)

            md = out_md.read_text(encoding="utf-8")
            self.assertIn("Excel-oracle mismatch report: **MISSING**", md)
            self.assertIn("| L2 | Calculate (Excel oracle) | MISSING |", md)
