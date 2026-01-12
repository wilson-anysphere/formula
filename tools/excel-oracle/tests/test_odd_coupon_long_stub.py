from __future__ import annotations

import json
import unittest
from pathlib import Path


class OddCouponLongStubCorpusTests(unittest.TestCase):
    def test_long_stub_cases_exist_in_canonical_corpus(self) -> None:
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
            for basis in ("basis0", "basis1"):
                matches = [
                    c
                    for c in cases
                    if isinstance(c, dict)
                    and has_all_tags(c, "odd_coupon", "long_stub", "financial", func, basis)
                ]
                self.assertTrue(
                    matches,
                    f"Missing long odd-coupon stub oracle coverage: tags odd_coupon+long_stub+{func}+{basis}",
                )

    def test_subset_corpus_formulas_are_in_canonical_corpus(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        subset_path = repo_root / "tools/excel-oracle/odd_coupon_long_stub_cases.json"
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")
        self.assertTrue(subset_path.is_file(), f"subset corpus not found at {subset_path}")

        cases_payload = json.loads(cases_path.read_text(encoding="utf-8"))
        cases = cases_payload.get("cases", [])
        self.assertIsInstance(cases, list, "cases.json top-level 'cases' must be an array")

        canonical_formulas: set[str] = {
            c["formula"] for c in cases if isinstance(c, dict) and isinstance(c.get("formula"), str)
        }

        subset_payload = json.loads(subset_path.read_text(encoding="utf-8"))
        subset_cases = subset_payload.get("cases", [])
        self.assertIsInstance(subset_cases, list, "subset corpus top-level 'cases' must be an array")
        self.assertGreaterEqual(
            len(subset_cases),
            8,
            "expected odd_coupon_long_stub_cases.json to contain the long-stub scenarios",
        )

        for c in subset_cases:
            if not isinstance(c, dict):
                continue
            formula = c.get("formula")
            self.assertIsInstance(formula, str)
            self.assertIn(
                formula,
                canonical_formulas,
                f"Subset corpus formula is missing from canonical cases.json: {formula}",
            )


if __name__ == "__main__":
    unittest.main()

