from __future__ import annotations

import unittest

from tools.corpus.triage import _scan_features


class TriageFeatureScanTests(unittest.TestCase):
    def test_scan_features_detects_cell_images(self) -> None:
        features = _scan_features(["xl/cellImages.xml"])
        self.assertIn("has_cell_images", features)
        self.assertTrue(features["has_cell_images"])

    def test_scan_features_detects_cell_images_with_leading_slash(self) -> None:
        # Some malformed XLSX archives store entry names like `/xl/cellImages.xml`. Feature scanning
        # should be tolerant and still detect the part.
        features = _scan_features(["/xl/cellImages.xml"])
        self.assertIn("has_cell_images", features)
        self.assertTrue(features["has_cell_images"])

    def test_scan_features_detects_cell_images_case_insensitively(self) -> None:
        features = _scan_features(["XL/CellImages1.XML"])
        self.assertIn("has_cell_images", features)
        self.assertTrue(features["has_cell_images"])

    def test_scan_features_detects_cell_images_in_folder_layout(self) -> None:
        features = _scan_features(["xl/cellImages/cellImages.xml"])
        self.assertIn("has_cell_images", features)
        self.assertTrue(features["has_cell_images"])

    def test_scan_features_detects_cell_images_numeric_suffix(self) -> None:
        features = _scan_features(["xl/cellimages1.xml"])
        self.assertIn("has_cell_images", features)
        self.assertTrue(features["has_cell_images"])

    def test_scan_features_cell_images_absent(self) -> None:
        features = _scan_features(["xl/workbook.xml"])
        self.assertIn("has_cell_images", features)
        self.assertFalse(features["has_cell_images"])


if __name__ == "__main__":
    unittest.main()
