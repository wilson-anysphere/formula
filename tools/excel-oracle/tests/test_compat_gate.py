from __future__ import annotations

import hashlib
import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path


class CompatGateDatasetSelectionTests(unittest.TestCase):
    def _load_compat_gate(self):
        compat_gate_py = Path(__file__).resolve().parents[1] / "compat_gate.py"
        self.assertTrue(compat_gate_py.is_file(), f"compat_gate.py not found at {compat_gate_py}")

        spec = importlib.util.spec_from_file_location("excel_oracle_compat_gate", compat_gate_py)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_default_expected_dataset_prefers_versioned_match_by_cases_sha(self) -> None:
        # Ensure compat_gate selects the expected dataset by matching the cases.json hash
        # (via the versioned filename suffix), not by lexicographic ordering.
        compat_gate = self._load_compat_gate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            old_cwd = Path.cwd()
            try:
                os.chdir(tmp_path)

                cases_path = tmp_path / "cases.json"
                cases_path.write_text('{"schemaVersion":1,"cases":[]}\n', encoding="utf-8", newline="\n")
                sha8 = hashlib.sha256(cases_path.read_bytes()).hexdigest()[:8]

                versioned_dir = tmp_path / "tests/compatibility/excel-oracle/datasets/versioned"
                versioned_dir.mkdir(parents=True, exist_ok=True)

                older = versioned_dir / f"excel-16.0-build-1-cases-{sha8}.json"
                newer = versioned_dir / f"excel-16.0-build-2-cases-{sha8}.json"
                older.write_text("{}", encoding="utf-8", newline="\n")
                newer.write_text("{}", encoding="utf-8", newline="\n")

                # Make sure `newer` has the newest mtime regardless of filesystem ordering.
                os.utime(older, (1, 1))
                os.utime(newer, (2, 2))

                selected = compat_gate._default_expected_dataset(cases_path=cases_path)
                self.assertEqual(selected.resolve(), newer.resolve())
            finally:
                os.chdir(old_cwd)

    def test_default_expected_dataset_falls_back_to_pinned(self) -> None:
        compat_gate = self._load_compat_gate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            old_cwd = Path.cwd()
            try:
                os.chdir(tmp_path)

                cases_path = tmp_path / "cases.json"
                cases_path.write_text('{"schemaVersion":1,"cases":[]}\n', encoding="utf-8", newline="\n")

                pinned = tmp_path / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
                pinned.parent.mkdir(parents=True, exist_ok=True)
                pinned.write_text("{}", encoding="utf-8", newline="\n")

                selected = compat_gate._default_expected_dataset(cases_path=cases_path)
                self.assertEqual(selected.resolve(), pinned.resolve())
            finally:
                os.chdir(old_cwd)


if __name__ == "__main__":
    unittest.main()
