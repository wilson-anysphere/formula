from __future__ import annotations

import json
import unittest
from pathlib import Path


class OddCouponBoundarySubsetTests(unittest.TestCase):
    def test_boundary_cases_exist_in_canonical_corpus(self) -> None:
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
                if isinstance(c, dict) and has_all_tags(c, "odd_coupon", "boundary", "financial", func)
            ]
            self.assertTrue(
                matches,
                f"Missing odd-coupon boundary oracle coverage for {func} (tags odd_coupon+boundary)",
            )

    def test_subset_case_ids_and_formulas_match_canonical_corpus(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        subset_path = repo_root / "tools/excel-oracle/odd_coupon_boundary_cases.json"
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")
        self.assertTrue(subset_path.is_file(), f"boundary subset corpus not found at {subset_path}")

        cases_payload = json.loads(cases_path.read_text(encoding="utf-8"))
        cases = cases_payload.get("cases", [])
        self.assertIsInstance(cases, list, "cases.json top-level 'cases' must be an array")

        canonical_by_id: dict[str, str] = {}
        for c in cases:
            if not isinstance(c, dict):
                continue
            cid = c.get("id")
            formula = c.get("formula")
            if isinstance(cid, str) and isinstance(formula, str):
                canonical_by_id[cid] = formula

        subset_payload = json.loads(subset_path.read_text(encoding="utf-8"))
        subset_cases = subset_payload.get("cases", [])
        self.assertIsInstance(subset_cases, list, "subset corpus top-level 'cases' must be an array")
        self.assertGreaterEqual(
            len(subset_cases),
            12,
            "expected odd_coupon_boundary_cases.json to contain boundary scenarios",
        )

        for c in subset_cases:
            if not isinstance(c, dict):
                continue
            cid = c.get("id")
            self.assertIsInstance(cid, str)
            self.assertIn(cid, canonical_by_id, f"Boundary subset case ID is missing from cases.json: {cid}")

            subset_formula = c.get("formula")
            self.assertIsInstance(subset_formula, str)
            self.assertEqual(
                subset_formula,
                canonical_by_id[cid],
                f"Boundary subset formula does not match canonical cases.json for caseId={cid}",
            )


if __name__ == "__main__":
    unittest.main()

