from __future__ import annotations

import unittest

from tools.corpus.triage import DEFAULT_DIFF_IGNORE


class CalcChainPolicyTests(unittest.TestCase):
    def test_calcchain_not_ignored_by_default(self) -> None:
        # CalcChain churn should be surfaced in corpus metrics/dashboards.
        self.assertNotIn("xl/calcChain.xml", DEFAULT_DIFF_IGNORE)


if __name__ == "__main__":
    unittest.main()

