from __future__ import annotations

import json
import unittest
from pathlib import Path


class LongStubSubsetSyncTests(unittest.TestCase):
    def test_long_stub_subset_file_is_subset_of_canonical_corpus(self) -> None:
        """
        The repo keeps a small convenience corpus for quickly pinning long-stub odd-coupon
        scenarios in real Excel:

          tools/excel-oracle/odd_coupon_long_stub_cases.json

        This test ensures it stays aligned with the canonical oracle corpus:

          tests/compatibility/excel-oracle/cases.json

        so the subset file never silently diverges (wrong formulas / missing tags).
        """

        repo_root = Path(__file__).resolve().parents[3]
        corpus_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        subset_path = repo_root / "tools/excel-oracle/odd_coupon_long_stub_cases.json"

        corpus = json.loads(corpus_path.read_text(encoding="utf-8"))
        subset = json.loads(subset_path.read_text(encoding="utf-8"))

        corpus_cases = [c for c in corpus.get("cases", []) if isinstance(c, dict)]
        subset_cases = [c for c in subset.get("cases", []) if isinstance(c, dict)]

        # Match subset cases by their exact formula; the subset file uses hand-written IDs.
        corpus_by_formula = {str(c.get("formula", "")): c for c in corpus_cases}

        for case in subset_cases:
            formula = str(case.get("formula", ""))
            self.assertTrue(formula, "subset case is missing formula")
            self.assertIn(formula, corpus_by_formula, f"subset formula not present in cases.json: {formula!r}")

            corpus_case = corpus_by_formula[formula]

            # Ensure the subset's tags are all present on the canonical case (canonical can have extras).
            subset_tags = set(case.get("tags", [])) if isinstance(case.get("tags"), list) else set()
            corpus_tags = set(corpus_case.get("tags", [])) if isinstance(corpus_case.get("tags"), list) else set()
            self.assertTrue(subset_tags.issubset(corpus_tags), f"tag drift for {formula!r}: {subset_tags} ⊄ {corpus_tags}")


if __name__ == "__main__":
    unittest.main()

