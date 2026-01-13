from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.corpus.promote_public import (
    _coerce_display_name,
    _write_public_fixture,
    extract_public_expectations,
    update_public_expectations_file,
    upsert_expectations_entry,
)


class PromotePublicExpectationsTests(unittest.TestCase):
    def test_write_public_fixture_refuses_overwrite_without_force(self) -> None:
        with tempfile.TemporaryDirectory(prefix="promote-public-fixture-test-") as td:
            tmp = Path(td)
            path = tmp / "book.xlsx.b64"
            _write_public_fixture(path, b"one", force=False)
            # Idempotent if bytes are identical (doesn't rewrite).
            before = path.read_bytes()
            _write_public_fixture(path, b"one", force=False)
            self.assertEqual(path.read_bytes(), before)

            with self.assertRaises(FileExistsError):
                _write_public_fixture(path, b"two", force=False)

            _write_public_fixture(path, b"two", force=True)
            self.assertNotEqual(path.read_bytes(), before)

    def test_coerce_display_name_appends_extension(self) -> None:
        self.assertEqual(_coerce_display_name("my-case", default_ext=".xlsx"), "my-case.xlsx")
        self.assertEqual(_coerce_display_name("my-case", default_ext=".xlsb"), "my-case.xlsb")

    def test_coerce_display_name_strips_b64_suffix(self) -> None:
        self.assertEqual(
            _coerce_display_name("my-case.xlsx.b64", default_ext=".xlsx"), "my-case.xlsx"
        )
        self.assertEqual(
            _coerce_display_name("my-case.xlsb.b64", default_ext=".xlsb"), "my-case.xlsb"
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

    def test_extract_public_expectations_requires_round_trip_ok_bool(self) -> None:
        with self.assertRaises(ValueError):
            extract_public_expectations({"result": {"open_ok": True, "round_trip_ok": None}})

    def test_extract_public_expectations_requires_diff_critical_int(self) -> None:
        with self.assertRaises(ValueError):
            extract_public_expectations(
                {"result": {"open_ok": True, "round_trip_ok": True, "diff_critical_count": "0"}}
            )

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

    def test_upsert_expectations_entry_is_idempotent_with_extra_existing_keys(self) -> None:
        expectations = {
            "book.xlsx": {
                "open_ok": True,
                "round_trip_ok": True,
                "diff_critical_count": 0,
                "diff_warning_count": 123,
            }
        }
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

            # If expectations contains extra keys, we should still treat it as up-to-date
            # as long as the required keys match.
            data = json.loads(expectations_path.read_text(encoding="utf-8"))
            data["book.xlsx"]["diff_warning_count"] = 123
            expectations_path.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
            before = expectations_path.read_text(encoding="utf-8")
            changed = update_public_expectations_file(
                expectations_path, workbook_name="book.xlsx", report=report, force=False
            )
            self.assertFalse(changed)
            self.assertEqual(expectations_path.read_text(encoding="utf-8"), before)

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
                {
                    "book.xlsx": {
                        "open_ok": True,
                        "round_trip_ok": False,
                        "diff_critical_count": 9,
                        "diff_warning_count": 123,
                    }
                },
            )


if __name__ == "__main__":
    unittest.main()
