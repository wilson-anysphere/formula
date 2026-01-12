from __future__ import annotations

import json
import unittest
from pathlib import Path


class OddCouponSubsetCorpusTests(unittest.TestCase):
    def test_odd_coupon_subset_corpora_match_canonical_cases(self) -> None:
        """Ensure the derived odd-coupon corpora stay in sync with cases.json.

        The JSON files under `tools/excel-oracle/odd_coupon_*.json` are intended to be small,
        hand-picked subsets of the canonical `tests/compatibility/excel-oracle/cases.json` corpus,
        used for quick Windows + Excel parity runs.

        They should therefore reference existing case IDs and preserve the same formulas/inputs.
        """

        repo_root = Path(__file__).resolve().parents[3]
        canonical_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        self.assertTrue(canonical_path.is_file(), f"cases.json not found at {canonical_path}")

        canonical = json.loads(canonical_path.read_text(encoding="utf-8"))
        canonical_cases = canonical.get("cases")
        self.assertIsInstance(canonical_cases, list, "cases.json 'cases' must be a list")

        canonical_by_id: dict[str, dict] = {}
        for case in canonical_cases:
            if not isinstance(case, dict):
                continue
            cid = case.get("id")
            if isinstance(cid, str):
                canonical_by_id[cid] = case

        subset_paths = [
            repo_root / "tools/excel-oracle/odd_coupon_boundary_cases.json",
            repo_root / "tools/excel-oracle/odd_coupon_invalid_schedule_cases.json",
            repo_root / "tools/excel-oracle/odd_coupon_basis4_cases.json",
            repo_root / "tools/excel-oracle/odd_coupon_long_stub_cases.json",
            repo_root / "tools/excel-oracle/odd_coupon_validation_cases.json",
        ]

        for subset_path in subset_paths:
            self.assertTrue(subset_path.is_file(), f"subset corpus not found at {subset_path}")

            subset = json.loads(subset_path.read_text(encoding="utf-8"))
            subset_cases = subset.get("cases")
            self.assertIsInstance(subset_cases, list, f"{subset_path} 'cases' must be a list")

            for subset_case in subset_cases:
                self.assertIsInstance(subset_case, dict, f"{subset_path} cases must be objects")
                cid = subset_case.get("id")
                self.assertIsInstance(cid, str, f"{subset_path} case id must be a string (got {cid!r})")

                canonical_case = canonical_by_id.get(cid)
                self.assertIsNotNone(
                    canonical_case,
                    f"{subset_path} references unknown caseId {cid!r} (not found in cases.json)",
                )
                assert canonical_case is not None

                # Preserve the executable shape of the case (formula + worksheet wiring + inputs).
                self.assertEqual(
                    subset_case.get("formula"),
                    canonical_case.get("formula"),
                    f"formula mismatch for caseId={cid} in {subset_path}",
                )
                self.assertEqual(
                    subset_case.get("outputCell"),
                    canonical_case.get("outputCell"),
                    f"outputCell mismatch for caseId={cid} in {subset_path}",
                )
                self.assertEqual(
                    subset_case.get("inputs"),
                    canonical_case.get("inputs"),
                    f"inputs mismatch for caseId={cid} in {subset_path}",
                )

                # Metadata is used for filtering/slicing and should stay aligned too.
                self.assertEqual(
                    subset_case.get("description"),
                    canonical_case.get("description"),
                    f"description mismatch for caseId={cid} in {subset_path}",
                )

                subset_tags = subset_case.get("tags", [])
                canonical_tags = canonical_case.get("tags", [])
                self.assertIsInstance(subset_tags, list)
                self.assertIsInstance(canonical_tags, list)
                self.assertEqual(
                    set(subset_tags),
                    set(canonical_tags),
                    f"tag mismatch for caseId={cid} in {subset_path}",
                )


if __name__ == "__main__":
    unittest.main()
