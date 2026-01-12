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

        by requiring it to contain *exactly* the cases tagged with BOTH:
          - odd_coupon
          - long_stub
        """

        repo_root = Path(__file__).resolve().parents[3]
        corpus_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        subset_path = repo_root / "tools/excel-oracle/odd_coupon_long_stub_cases.json"

        corpus = json.loads(corpus_path.read_text(encoding="utf-8"))
        subset = json.loads(subset_path.read_text(encoding="utf-8"))

        corpus_cases = [c for c in corpus.get("cases", []) if isinstance(c, dict)]
        subset_cases = [c for c in subset.get("cases", []) if isinstance(c, dict)]

        long_stub_cases = [
            c
            for c in corpus_cases
            if isinstance(c.get("tags"), list)
            and "odd_coupon" in c["tags"]
            and "long_stub" in c["tags"]
        ]
        self.assertTrue(long_stub_cases, "Expected at least one odd_coupon+long_stub case in cases.json")

        expected_ids = {c.get("id") for c in long_stub_cases if isinstance(c.get("id"), str)}
        subset_ids = {c.get("id") for c in subset_cases if isinstance(c.get("id"), str)}
        self.assertEqual(
            subset_ids,
            expected_ids,
            "Long-stub subset file must contain exactly the odd_coupon+long_stub case IDs from cases.json",
        )

        corpus_by_id = {c["id"]: c for c in corpus_cases if isinstance(c.get("id"), str)}
        for case in subset_cases:
            cid = case.get("id")
            self.assertIsInstance(cid, str, "subset case is missing id")
            self.assertIn(cid, corpus_by_id, f"subset caseId not present in cases.json: {cid!r}")

            corpus_case = corpus_by_id[cid]

            # Ensure the subset case matches the canonical one, so any drift is obvious.
            self.assertEqual(case.get("formula"), corpus_case.get("formula"), f"formula drift for {cid!r}")
            self.assertEqual(case.get("outputCell"), corpus_case.get("outputCell"), f"outputCell drift for {cid!r}")
            self.assertEqual(case.get("inputs"), corpus_case.get("inputs"), f"inputs drift for {cid!r}")
            self.assertEqual(
                case.get("description"),
                corpus_case.get("description"),
                f"description drift for {cid!r}",
            )

            subset_tags = set(case.get("tags", [])) if isinstance(case.get("tags"), list) else set()
            corpus_tags = set(corpus_case.get("tags", [])) if isinstance(corpus_case.get("tags"), list) else set()
            self.assertIn("odd_coupon", subset_tags, f"subset case is missing odd_coupon tag: {cid!r}")
            self.assertIn("long_stub", subset_tags, f"subset case is missing long_stub tag: {cid!r}")
            self.assertTrue(
                subset_tags.issubset(corpus_tags),
                f"tag drift for {cid!r}: {subset_tags} ⊄ {corpus_tags}",
            )


if __name__ == "__main__":
    unittest.main()
