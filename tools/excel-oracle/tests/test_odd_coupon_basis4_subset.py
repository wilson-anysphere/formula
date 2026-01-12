from __future__ import annotations

import json
import unittest
from pathlib import Path


class OddCouponBasis4SubsetCorpusTests(unittest.TestCase):
    def test_basis4_cases_exist_in_canonical_corpus(self) -> None:
        """Ensure we keep at least one odd-coupon basis=4 oracle case per ODDF*/ODDL* function."""

        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")

        cases_payload = json.loads(cases_path.read_text(encoding="utf-8"))
        cases = cases_payload.get("cases", [])
        self.assertIsInstance(cases, list, "cases.json top-level 'cases' must be an array")

        def has_all_tags(case: dict, *tags: str) -> bool:
            tag_list = case.get("tags", [])
            if not isinstance(tag_list, list):
                return False
            return all(t in tag_list for t in tags)

        for func in ("ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD"):
            matches = [
                c
                for c in cases
                if isinstance(c, dict)
                and has_all_tags(c, "financial", "odd_coupon", "basis4", func)
            ]
            self.assertTrue(
                matches,
                f"Missing odd-coupon basis=4 oracle coverage: tags financial+odd_coupon+basis4+{func}",
            )

    def test_subset_corpus_ids_and_formulas_match_canonical(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        subset_path = repo_root / "tools/excel-oracle/odd_coupon_basis4_cases.json"
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")
        self.assertTrue(subset_path.is_file(), f"subset corpus not found at {subset_path}")

        cases_payload = json.loads(cases_path.read_text(encoding="utf-8"))
        cases = cases_payload.get("cases", [])
        self.assertIsInstance(cases, list, "cases.json top-level 'cases' must be an array")

        basis4_cases = [
            c
            for c in cases
            if isinstance(c, dict)
            and isinstance(c.get("tags"), list)
            and "odd_coupon" in c["tags"]
            and "basis4" in c["tags"]
        ]
        self.assertTrue(basis4_cases, "Expected at least one odd_coupon+basis4 case in cases.json")

        canonical_by_id: dict[str, dict] = {
            c["id"]: c for c in cases if isinstance(c, dict) and isinstance(c.get("id"), str)
        }

        subset_payload = json.loads(subset_path.read_text(encoding="utf-8"))
        subset_cases = subset_payload.get("cases", [])
        self.assertIsInstance(subset_cases, list, "subset corpus top-level 'cases' must be an array")

        expected_ids = {c.get("id") for c in basis4_cases if isinstance(c.get("id"), str)}
        subset_ids = {
            c.get("id") for c in subset_cases if isinstance(c, dict) and isinstance(c.get("id"), str)
        }
        self.assertEqual(
            subset_ids,
            expected_ids,
            "Basis4 subset file must contain exactly the odd_coupon+basis4 case IDs from cases.json",
        )

        for c in subset_cases:
            if not isinstance(c, dict):
                continue
            cid = c.get("id")
            self.assertIsInstance(cid, str)
            self.assertIn(
                cid,
                canonical_by_id,
                f"Subset corpus case ID is missing from canonical cases.json: {cid}",
            )

            canonical = canonical_by_id[cid]
            self.assertEqual(
                c.get("formula"),
                canonical.get("formula"),
                f"Subset corpus formula drift for caseId {cid}",
            )

            subset_tags = set(c.get("tags", [])) if isinstance(c.get("tags"), list) else set()
            canonical_tags = (
                set(canonical.get("tags", [])) if isinstance(canonical.get("tags"), list) else set()
            )
            self.assertTrue(
                subset_tags.issubset(canonical_tags),
                f"Subset corpus tag drift for caseId {cid}: {sorted(subset_tags)} ⊄ {sorted(canonical_tags)}",
            )


if __name__ == "__main__":
    unittest.main()

