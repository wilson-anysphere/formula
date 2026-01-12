from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
import unittest
from pathlib import Path


class PinnedDatasetSyncTests(unittest.TestCase):
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

    @staticmethod
    def _sanitize_fragment(text: str) -> str:
        # Keep in sync with tools/excel-oracle/pin_dataset.py::_sanitize_fragment.
        import re

        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", text.strip())
        safe = re.sub(r"_+", "_", safe).strip("_")
        return safe or "unknown"

    def test_pinned_dataset_hash_matches_cases_json(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        pinned_path = repo_root / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"

        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")
        self.assertTrue(pinned_path.is_file(), f"pinned dataset not found at {pinned_path}")

        cases_sha = hashlib.sha256(cases_path.read_bytes()).hexdigest()
        pinned = json.loads(pinned_path.read_text(encoding="utf-8"))

        case_set = pinned.get("caseSet")
        self.assertIsInstance(case_set, dict, "Pinned dataset is missing caseSet metadata")

        self.assertEqual(case_set.get("sha256"), cases_sha, "Pinned dataset caseSet.sha256 must match cases.json")

        results = pinned.get("results", [])
        self.assertIsInstance(results, list, "Pinned dataset results must be an array")
        self.assertEqual(case_set.get("count"), len(results), "Pinned dataset caseSet.count must match results length")

        # Ensure the pinned dataset covers the full corpus.
        #
        # This file (`excel-oracle.pinned.json`) is the default fallback oracle dataset used by
        # `tools/excel-oracle/compat_gate.py` when no versioned Excel dataset is available. If it
        # becomes a strict subset of cases.json (e.g. cases are added but the pinned dataset isn't
        # regenerated), compare.py will fail with "missing-expected" mismatches.
        #
        # Keep this strict so adding new deterministic functions/cases forces a corresponding pin
        # update.
        cases_payload = json.loads(cases_path.read_text(encoding="utf-8"))
        case_ids = {
            c.get("id")
            for c in cases_payload.get("cases", [])
            if isinstance(c, dict) and isinstance(c.get("id"), str)
        }
        result_ids = {r.get("caseId") for r in results if isinstance(r, dict) and isinstance(r.get("caseId"), str)}
        self.assertEqual(
            result_ids,
            case_ids,
            "Pinned dataset must include results for every case ID in cases.json",
        )
        self.assertEqual(
            len(results),
            len(case_ids),
            "Pinned dataset results must have exactly one entry per case (no duplicates, no omissions)",
        )

    def test_pinned_dataset_is_in_sync_with_matching_versioned_copy(self) -> None:
        # `tools/excel-oracle/pin_dataset.py` writes:
        # - a pinned dataset at `datasets/excel-oracle.pinned.json`, and
        # - a versioned copy in `datasets/versioned/` with a filename derived from the pinned
        #   dataset's Excel version/build and cases.json SHA-256 suffix.
        #
        # The compatibility gate (`tools/excel-oracle/compat_gate.py`) prefers versioned datasets
        # when available. If the pinned dataset is updated without updating its corresponding
        # versioned copy, CI can regress with confusing mismatches.
        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        pinned_path = repo_root / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
        versioned_dir = repo_root / "tests/compatibility/excel-oracle/datasets/versioned"

        pinned = json.loads(pinned_path.read_text(encoding="utf-8"))
        source = pinned.get("source", {})
        case_set = pinned.get("caseSet", {})

        self.assertIsInstance(source, dict)
        self.assertIsInstance(case_set, dict)

        cases_sha8 = hashlib.sha256(cases_path.read_bytes()).hexdigest()[:8]
        excel_version = self._sanitize_fragment(str(source.get("version", "unknown")))
        excel_build = self._sanitize_fragment(str(source.get("build", "unknown")))
        expected_name = f"excel-{excel_version}-build-{excel_build}-cases-{cases_sha8}.json"
        versioned_path = versioned_dir / expected_name

        # Only enforce the sync check when the corresponding versioned file exists.
        # (The versioned directory is optional; compat_gate falls back to the pinned dataset when
        # there is no matching versioned dataset for the current cases.json hash.)
        if not versioned_path.is_file():
            return

        versioned = json.loads(versioned_path.read_text(encoding="utf-8"))
        self.assertEqual(
            versioned,
            pinned,
            f"Versioned dataset {versioned_path} must be an exact copy of the pinned dataset {pinned_path}",
        )

        # Ensure compat_gate selection is consistent with the pinned dataset when it selects the
        # pinned dataset's version/build.
        compat_gate = self._load_compat_gate()
        selected_path = compat_gate._default_expected_dataset(cases_path=cases_path)
        if selected_path.resolve() == versioned_path.resolve():
            selected = json.loads(selected_path.read_text(encoding="utf-8"))
            self.assertEqual(selected, pinned)


if __name__ == "__main__":
    unittest.main()
