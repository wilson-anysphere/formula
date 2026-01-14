from __future__ import annotations

import base64
import contextlib
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import tools.corpus.promote_public as promote_mod
from tools.corpus.util import sha256_hex


def _make_minimal_xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</workbook>
""",
        )
    return buf.getvalue()


class PromotePublicCLITests(unittest.TestCase):
    def test_main_defaults_to_hash_name_outside_public_dir(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_bytes = _make_minimal_xlsx()
            input_path.write_bytes(input_bytes)

            expected_name = f"workbook-{sha256_hex(input_bytes)[:16]}.xlsx"

            original_run_public_triage = promote_mod._run_public_triage
            try:
                promote_mod._run_public_triage = (  # type: ignore[assignment]
                    lambda wb, *, diff_limit=25, recalc=False, render_smoke=False: {
                        "sha256": sha256_hex(wb.data),
                        "result": {
                            "open_ok": True,
                            "round_trip_ok": True,
                            "diff_critical_count": 0,
                        },
                    }
                )

                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)

                fixture_path = public_dir / f"{expected_name}.b64"
                self.assertTrue(fixture_path.exists())

                expectations = json.loads((public_dir / "expectations.json").read_text(encoding="utf-8"))
                self.assertIn(expected_name, expectations)
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_dry_run_does_not_write_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_bytes = _make_minimal_xlsx()
            input_path.write_bytes(input_bytes)

            original_run_public_triage = promote_mod._run_public_triage
            try:
                promote_mod._run_public_triage = (  # type: ignore[assignment]
                    lambda wb, *, diff_limit=25, recalc=False, render_smoke=False: {
                        "sha256": sha256_hex(wb.data),
                        "result": {
                            "open_ok": True,
                            "round_trip_ok": True,
                            "diff_critical_count": 0,
                        },
                    }
                )

                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                            "--dry-run",
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)
                out = json.loads(stdout.getvalue())
                self.assertTrue(out.get("dry_run"))
                self.assertTrue(out.get("fixture_changed"))
                self.assertTrue(out.get("expectations_changed"))
                self.assertIsNone(out.get("triage_report"))

                # No writes.
                self.assertFalse(public_dir.exists())
                self.assertFalse(triage_out.exists())
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_dry_run_reports_need_force_on_existing_entries(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            public_dir.mkdir(parents=True, exist_ok=True)

            # Existing fixture/expectations that will disagree with the new triage result.
            old_bytes = _make_minimal_xlsx()
            new_bytes = _make_minimal_xlsx()
            # Make bytes differ while still being a valid XLSX zip.
            new_bytes = new_bytes + b"\n"

            fixture_path = public_dir / "case.xlsx.b64"
            fixture_path.write_bytes(base64.encodebytes(old_bytes))
            expectations_path = public_dir / "expectations.json"
            expectations_path.write_text(
                json.dumps(
                    {"case.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}},
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )

            input_path = tmp / "input.xlsx"
            input_path.write_bytes(new_bytes)

            fixture_before = fixture_path.read_bytes()
            expectations_before = expectations_path.read_bytes()

            original_run_public_triage = promote_mod._run_public_triage
            try:
                promote_mod._run_public_triage = (  # type: ignore[assignment]
                    lambda wb, *, diff_limit=25, recalc=False, render_smoke=False: {
                        "sha256": sha256_hex(wb.data),
                        "result": {
                            "open_ok": True,
                            "round_trip_ok": True,
                            "diff_critical_count": 1,
                        },
                    }
                )

                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--name",
                            "case.xlsx",
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                            "--dry-run",
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)
                out = json.loads(stdout.getvalue())
                self.assertTrue(out.get("dry_run"))
                self.assertTrue(out.get("fixture_changed"))
                self.assertTrue(out.get("expectations_changed"))
                needs_force = out.get("needs_force") or {}
                self.assertEqual(needs_force.get("fixture"), True)
                self.assertEqual(needs_force.get("expectations"), True)

                # Ensure we didn't mutate any tracked artifacts.
                self.assertEqual(fixture_path.read_bytes(), fixture_before)
                self.assertEqual(expectations_path.read_bytes(), expectations_before)
                self.assertFalse(triage_out.exists())
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_skips_triage_for_existing_public_fixture(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            public_dir.mkdir(parents=True, exist_ok=True)

            xlsx_bytes = _make_minimal_xlsx()
            fixture_path = public_dir / "case.xlsx.b64"
            fixture_path.write_bytes(base64.encodebytes(xlsx_bytes))

            (public_dir / "expectations.json").write_text(
                json.dumps(
                    {"case.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}},
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )

            called = {"triage": 0}
            original_run_public_triage = promote_mod._run_public_triage
            try:
                def _triage_should_not_run(*args, **kwargs):  # type: ignore[no-untyped-def]
                    called["triage"] += 1
                    raise RuntimeError("triage should not run")

                promote_mod._run_public_triage = _triage_should_not_run  # type: ignore[assignment]

                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(fixture_path),
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)
                self.assertEqual(called["triage"], 0)
                self.assertIn("already_promoted", stdout.getvalue())
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_idempotent_for_existing_fixture_without_rust(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_bytes = _make_minimal_xlsx()
            input_path.write_bytes(input_bytes)

            original_run_public_triage = promote_mod._run_public_triage
            try:
                # Avoid Rust build by stubbing triage.
                def _fake_triage(  # type: ignore[no-untyped-def]
                    wb, *, diff_limit: int = 25, recalc: bool = False, render_smoke: bool = False
                ):
                    return {
                        "sha256": sha256_hex(wb.data),
                        "result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0},
                    }

                promote_mod._run_public_triage = _fake_triage  # type: ignore[assignment]

                def _run(argv: list[str]) -> int:
                    stdout = io.StringIO()
                    with contextlib.redirect_stdout(stdout):
                        with mock.patch.object(sys, "argv", argv):
                            return promote_mod.main()

                argv = [
                    "promote_public.py",
                    "--input",
                    str(input_path),
                    "--name",
                    "case",
                    "--public-dir",
                    str(public_dir),
                    "--triage-out",
                    str(triage_out),
                    "--diff-limit",
                    "10",
                ]

                rc = _run(argv)
                self.assertEqual(rc, 0)

                fixture_path = public_dir / "case.xlsx.b64"
                expectations_path = public_dir / "expectations.json"
                self.assertTrue(fixture_path.exists())
                self.assertTrue(expectations_path.exists())

                fixture_before = fixture_path.read_bytes()
                expectations_before = expectations_path.read_bytes()

                # Second run should not change the tracked artifacts.
                rc = _run(argv)
                self.assertEqual(rc, 0)
                self.assertEqual(fixture_path.read_bytes(), fixture_before)
                self.assertEqual(expectations_path.read_bytes(), expectations_before)
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_refuses_to_overwrite_expectations_without_force(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_path.write_bytes(_make_minimal_xlsx())

            original_run_public_triage = promote_mod._run_public_triage
            try:
                # First run: diff_critical_count=0
                promote_mod._run_public_triage = (  # type: ignore[assignment]
                    lambda wb, *, diff_limit=25, recalc=False, render_smoke=False: {
                        "sha256": sha256_hex(wb.data),
                        "result": {
                            "open_ok": True,
                            "round_trip_ok": True,
                            "diff_critical_count": 0,
                        },
                    }
                )
                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--name",
                            "case",
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)

                # Second run: new expectations, should refuse without --force.
                promote_mod._run_public_triage = (  # type: ignore[assignment]
                    lambda wb, *, diff_limit=25, recalc=False, render_smoke=False: {
                        "sha256": sha256_hex(wb.data),
                        "result": {
                            "open_ok": True,
                            "round_trip_ok": True,
                            "diff_critical_count": 1,
                        },
                    }
                )
                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--name",
                            "case",
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 1)

                # With --force it should overwrite.
                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(input_path),
                            "--name",
                            "case",
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                            "--force",
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)

                expectations_path = public_dir / "expectations.json"
                data = json.loads(expectations_path.read_text(encoding="utf-8"))
                self.assertEqual(data["case.xlsx"]["diff_critical_count"], 1)
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_rejects_name_with_path_separators(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_path.write_bytes(_make_minimal_xlsx())

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                with mock.patch.object(
                    sys,
                    "argv",
                    [
                        "promote_public.py",
                        "--input",
                        str(input_path),
                        "--name",
                        "bad/name",
                        "--public-dir",
                        str(public_dir),
                        "--triage-out",
                        str(triage_out),
                    ],
                ):
                    rc = promote_mod.main()
            self.assertEqual(rc, 1)

    def test_main_fails_gracefully_on_non_xlsx_input(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_path.write_text("not a zip", encoding="utf-8")

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                with mock.patch.object(
                    sys,
                    "argv",
                    [
                        "promote_public.py",
                        "--input",
                        str(input_path),
                        "--name",
                        "case",
                        "--public-dir",
                        str(public_dir),
                        "--triage-out",
                        str(triage_out),
                    ],
                ):
                    rc = promote_mod.main()
            self.assertEqual(rc, 1)
            self.assertIn("Leak scan failed", stdout.getvalue())

    def test_main_requires_confirm_for_xlsb(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsb"
            input_path.write_bytes(b"not really xlsb")

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                with mock.patch.object(
                    sys,
                    "argv",
                    [
                        "promote_public.py",
                        "--input",
                        str(input_path),
                        "--name",
                        "case",
                        "--public-dir",
                        str(public_dir),
                        "--triage-out",
                        str(triage_out),
                    ],
                ):
                    rc = promote_mod.main()
            self.assertEqual(rc, 1)
            self.assertIn("XLSB leak scanning is not supported", stdout.getvalue())

    def test_main_skips_xlsb_fixture_without_confirm(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            public_dir.mkdir(parents=True, exist_ok=True)

            fixture_path = public_dir / "case.xlsb.b64"
            fixture_path.write_bytes(base64.encodebytes(b"dummy-xlsb"))
            (public_dir / "expectations.json").write_text(
                json.dumps(
                    {"case.xlsb": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}},
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )

            called = {"triage": 0}
            original_run_public_triage = promote_mod._run_public_triage
            try:
                def _triage_should_not_run(*args, **kwargs):  # type: ignore[no-untyped-def]
                    called["triage"] += 1
                    raise RuntimeError("triage should not run")

                promote_mod._run_public_triage = _triage_should_not_run  # type: ignore[assignment]

                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    with mock.patch.object(
                        sys,
                        "argv",
                        [
                            "promote_public.py",
                            "--input",
                            str(fixture_path),
                            "--public-dir",
                            str(public_dir),
                            "--triage-out",
                            str(triage_out),
                        ],
                    ):
                        rc = promote_mod.main()
                self.assertEqual(rc, 0)
                self.assertEqual(called["triage"], 0)
                self.assertIn("already_promoted", stdout.getvalue())
            finally:
                promote_mod._run_public_triage = original_run_public_triage  # type: ignore[assignment]

    def test_main_rejects_name_with_mismatched_extension(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-cli-test-") as td:
            tmp = Path(td)
            public_dir = tmp / "public"
            triage_out = tmp / "triage"
            input_path = tmp / "input.xlsx"
            input_path.write_bytes(_make_minimal_xlsx())

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                with mock.patch.object(
                    sys,
                    "argv",
                    [
                        "promote_public.py",
                        "--input",
                        str(input_path),
                        "--name",
                        "case.xlsm",
                        "--public-dir",
                        str(public_dir),
                        "--triage-out",
                        str(triage_out),
                    ],
                ):
                    rc = promote_mod.main()
            self.assertEqual(rc, 1)
            self.assertIn("--name extension must match input extension", stdout.getvalue())


if __name__ == "__main__":
    unittest.main()
