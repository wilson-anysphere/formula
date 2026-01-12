from __future__ import annotations

import hashlib
import json
import unittest
from pathlib import Path


class PinnedDatasetSyncTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
