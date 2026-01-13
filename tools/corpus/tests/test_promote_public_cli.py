from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

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
                def _fake_triage(wb, *, diff_limit: int = 25):  # type: ignore[no-untyped-def]
                    return {
                        "sha256": sha256_hex(wb.data),
                        "result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0},
                    }

                promote_mod._run_public_triage = _fake_triage  # type: ignore[assignment]

                def _run(argv: list[str]) -> int:
                    stdout = io.StringIO()
                    with contextlib.redirect_stdout(stdout):
                        with unittest.mock.patch.object(sys, "argv", argv):
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
                    lambda wb, *, diff_limit=25: {
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
                    with unittest.mock.patch.object(
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
                    lambda wb, *, diff_limit=25: {
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
                    with unittest.mock.patch.object(
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
                    with unittest.mock.patch.object(
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
                with unittest.mock.patch.object(
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


if __name__ == "__main__":
    unittest.main()
