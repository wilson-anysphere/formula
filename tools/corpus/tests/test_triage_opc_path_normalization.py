from __future__ import annotations

import unittest

from tools.corpus.triage import _normalize_opc_path


class TriageOpcPathNormalizationTests(unittest.TestCase):
    def test_normalize_opc_path_resolves_dot_dot_segments(self) -> None:
        self.assertEqual(
            _normalize_opc_path("xl/drawings/_rels/../media/image1.png"),
            "xl/drawings/media/image1.png",
        )

    def test_normalize_opc_path_handles_backslashes_and_leading_slash(self) -> None:
        self.assertEqual(
            _normalize_opc_path("/xl\\drawings\\_rels\\..\\media\\image1.png"),
            "xl/drawings/media/image1.png",
        )

    def test_normalize_opc_path_does_not_escape_package_root(self) -> None:
        self.assertEqual(_normalize_opc_path("xl/../../foo/bar.xml"), "foo/bar.xml")


if __name__ == "__main__":
    unittest.main()

