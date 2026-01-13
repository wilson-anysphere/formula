from __future__ import annotations

import io
import json
import unittest
import zipfile
from pathlib import Path

from tools.corpus.triage import _step_failed, triage_workbook
from tools.corpus.util import WorkbookInput, sha256_hex


def _make_empty_zip() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED):
        pass
    return buf.getvalue()


class TriagePrivacyTests(unittest.TestCase):
    def test_step_failed_hashes_exception_string(self) -> None:
        err = ValueError("Sensitive /path/to/secret.xlsx")
        step = _step_failed(0.0, err)
        expected = f"sha256={sha256_hex(str(err).encode('utf-8'))}"
        self.assertEqual(step.error, expected)
        self.assertNotIn("secret.xlsx", step.error or "")

    def test_triage_workbook_hashes_triage_error_step(self) -> None:
        import tools.corpus.triage as triage_mod

        secret_msg = "secret file path /home/alice/private.xlsx"
        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            def _boom(*args, **kwargs):  # type: ignore[no-untyped-def]
                raise RuntimeError(secret_msg)

            triage_mod._run_rust_triage = _boom  # type: ignore[assignment]

            wb = WorkbookInput(display_name="book.xlsx", data=_make_empty_zip())
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        self.assertEqual(report.get("failure_category"), "triage_error")
        load_step = report["steps"]["load"]
        self.assertEqual(load_step["status"], "failed")
        expected = f"sha256={sha256_hex(secret_msg.encode('utf-8'))}"
        self.assertEqual(load_step["error"], expected)
        self.assertNotIn(secret_msg, json.dumps(report))

    def test_triage_workbook_hashes_feature_scan_errors(self) -> None:
        import tools.corpus.triage as triage_mod

        # Capture the raw exception message that zipfile emits for invalid archives so we can
        # assert that triage hashes it deterministically.
        invalid_zip = b"not a zip file"
        try:
            with zipfile.ZipFile(io.BytesIO(invalid_zip), "r"):
                pass
        except Exception as e:  # noqa: BLE001
            raw_msg = str(e)

        original_run_rust_triage = triage_mod._run_rust_triage
        try:
            triage_mod._run_rust_triage = lambda *args, **kwargs: {  # type: ignore[assignment]
                "steps": {},
                "result": {"open_ok": True, "round_trip_ok": True},
            }

            wb = WorkbookInput(display_name="book.xlsx", data=invalid_zip)
            report = triage_workbook(
                wb,
                rust_exe=Path("noop"),
                diff_ignore=set(),
                diff_limit=0,
                recalc=False,
                render_smoke=False,
            )
        finally:
            triage_mod._run_rust_triage = original_run_rust_triage  # type: ignore[assignment]

        expected = f"sha256={sha256_hex(raw_msg.encode('utf-8'))}"
        self.assertEqual(report["features_error"], expected)
        self.assertNotIn(raw_msg, json.dumps(report))


if __name__ == "__main__":
    unittest.main()

