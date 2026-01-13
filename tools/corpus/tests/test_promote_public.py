from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.corpus.promote_public import (
    _coerce_display_name,
    extract_public_expectations,
    update_public_expectations_file,
    upsert_expectations_entry,
)


class PromotePublicExpectationsTests(unittest.TestCase):
    def test_coerce_display_name_appends_extension(self) -> None:
        self.assertEqual(_coerce_display_name("my-case", default_ext=".xlsx"), "my-case.xlsx")

    def test_coerce_display_name_strips_b64_suffix(self) -> None:
        self.assertEqual(
            _coerce_display_name("my-case.xlsx.b64", default_ext=".xlsx"), "my-case.xlsx"
        )

    def test_coerce_display_name_rejects_path_separators(self) -> None:
        with self.assertRaises(ValueError):
            _coerce_display_name("foo/bar", default_ext=".xlsx")
        with self.assertRaises(ValueError):
            _coerce_display_name("foo\\bar", default_ext=".xlsx")

    def test_extract_public_expectations_happy_path(self) -> None:
        report = {"result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 7}}
        entry = extract_public_expectations(report)
        self.assertEqual(
            entry,
            {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 7},
        )

    def test_extract_public_expectations_requires_open_ok_true(self) -> None:
        with self.assertRaises(ValueError):
            extract_public_expectations({"result": {"open_ok": False}})

    def test_upsert_expectations_entry_idempotent(self) -> None:
        expectations = {"book.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}}
        updated, changed = upsert_expectations_entry(
            expectations=expectations,
            workbook_name="book.xlsx",
            entry={"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0},
            force=False,
        )
        self.assertFalse(changed)
        self.assertEqual(updated, expectations)

    def test_upsert_expectations_entry_refuses_overwrite_without_force(self) -> None:
        expectations = {"book.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 0}}
        with self.assertRaises(FileExistsError):
            upsert_expectations_entry(
                expectations=expectations,
                workbook_name="book.xlsx",
                entry={"open_ok": True, "round_trip_ok": False, "diff_critical_count": 1},
                force=False,
            )

    def test_update_public_expectations_file_writes_and_respects_force(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-test-") as td:
            tmp = Path(td)
            expectations_path = tmp / "expectations.json"
            report = {"result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 2}}

            # New file: should be written.
            changed = update_public_expectations_file(
                expectations_path, workbook_name="book.xlsx", report=report, force=False
            )
            self.assertTrue(changed)
            data = json.loads(expectations_path.read_text(encoding="utf-8"))
            self.assertEqual(
                data,
                {"book.xlsx": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": 2}},
            )

            # Idempotent: same report should not change the file.
            changed = update_public_expectations_file(
                expectations_path, workbook_name="book.xlsx", report=report, force=False
            )
            self.assertFalse(changed)

            # Different expectations without --force should fail.
            report2 = {"result": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 9}}
            with self.assertRaises(FileExistsError):
                update_public_expectations_file(
                    expectations_path, workbook_name="book.xlsx", report=report2, force=False
                )

            # With --force it should overwrite.
            changed = update_public_expectations_file(
                expectations_path, workbook_name="book.xlsx", report=report2, force=True
            )
            self.assertTrue(changed)
            data = json.loads(expectations_path.read_text(encoding="utf-8"))
            self.assertEqual(
                data,
                {"book.xlsx": {"open_ok": True, "round_trip_ok": False, "diff_critical_count": 9}},
            )


if __name__ == "__main__":
    unittest.main()
