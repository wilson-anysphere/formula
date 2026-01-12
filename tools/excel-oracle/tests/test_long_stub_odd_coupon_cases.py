from __future__ import annotations

import json
import unittest
from pathlib import Path


class LongStubOddCouponCasesTests(unittest.TestCase):
    def test_long_stub_odd_coupon_cases_are_in_canonical_corpus(self) -> None:
        """
        Regression test for Task 85 ("odd-coupon long stubs"):

        Ensure the curated `cases.json` corpus includes long-stub odd-coupon bond
        scenarios (ODDF*/ODDL*) so they are exercised by the Excel-oracle pipeline.
        """

        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        payload = json.loads(cases_path.read_text(encoding="utf-8"))

        cases = [c for c in payload.get("cases", []) if isinstance(c, dict)]
        long_stub_cases = [
            c
            for c in cases
            if isinstance(c.get("tags"), list)
            and "odd_coupon" in c["tags"]
            and "long_stub" in c["tags"]
        ]

        self.assertGreaterEqual(
            len(long_stub_cases),
            8,
            "Expected at least 8 long-stub odd-coupon cases (4 functions x basis0/basis1).",
        )

        for func in ["ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD"]:
            for basis_tag in ["basis0", "basis1"]:
                matches = [
                    c
                    for c in long_stub_cases
                    if isinstance(c.get("tags"), list)
                    and func in c["tags"]
                    and basis_tag in c["tags"]
                ]
                self.assertTrue(matches, f"Missing {func} long_stub {basis_tag} case in cases.json")

                # Prefer deterministic DATE(...) inputs (no volatile TODAY/NOW).
                for case in matches:
                    formula = str(case.get("formula", ""))
                    self.assertNotIn("TODAY(", formula.upper(), f"Volatile TODAY() in formula: {formula!r}")
                    self.assertNotIn("NOW(", formula.upper(), f"Volatile NOW() in formula: {formula!r}")
                    self.assertIn("DATE(", formula.upper(), f"Expected fixed DATE() inputs in formula: {formula!r}")


if __name__ == "__main__":
    unittest.main()

