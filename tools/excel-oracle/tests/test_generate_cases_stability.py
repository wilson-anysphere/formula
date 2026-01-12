from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
import unittest
from pathlib import Path


class GenerateCasesStabilityTests(unittest.TestCase):
    def test_generate_cases_matches_committed_cases_json(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        generator_path = repo_root / "tools/excel-oracle/generate_cases.py"
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"

        self.assertTrue(generator_path.is_file(), f"generate_cases.py not found at {generator_path}")
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")

        # `generate_cases.py` imports `case_generators` as a top-level package, so we need to
        # ensure the script directory is importable.
        sys_path_before = list(sys.path)
        sys.path.insert(0, str(generator_path.parent))
        try:
            spec = importlib.util.spec_from_file_location("_excel_oracle_generate_cases", generator_path)
            self.assertIsNotNone(spec)
            self.assertIsNotNone(spec.loader)
            module = importlib.util.module_from_spec(spec)
            # Some stdlib helpers (e.g. `dataclasses`) expect the module to be registered.
            sys.modules[spec.name] = module
            spec.loader.exec_module(module)  # type: ignore[union-attr]

            payload = module.generate_cases()
            # Serialize with the exact same formatting as `generate_cases.py` uses when writing.
            generated_bytes = (json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=False) + "\n").encode("utf-8")
        finally:
            sys.path[:] = sys_path_before

        expected_bytes = cases_path.read_bytes()

        expected_payload = json.loads(expected_bytes)
        self.assertEqual(
            len(payload.get("cases", [])),
            len(expected_payload.get("cases", [])),
            "Generated case count differs from committed cases.json",
        )

        generated_sha = hashlib.sha256(generated_bytes).hexdigest()
        expected_sha = hashlib.sha256(expected_bytes).hexdigest()
        self.assertEqual(generated_sha, expected_sha, "Generated cases.json does not match committed file (corpus churn)")


if __name__ == "__main__":
    unittest.main()
