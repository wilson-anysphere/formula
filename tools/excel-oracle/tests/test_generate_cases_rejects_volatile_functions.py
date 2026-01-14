from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path


class GenerateCasesVolatileValidationTests(unittest.TestCase):
    def _load_generator(self):
        repo_root = Path(__file__).resolve().parents[3]
        generator_path = repo_root / "tools/excel-oracle/generate_cases.py"
        self.assertTrue(generator_path.is_file(), f"generate_cases.py not found at {generator_path}")

        # `generate_cases.py` imports `case_generators` as a top-level package, so ensure the script
        # directory is importable.
        sys_path_before = list(sys.path)
        sys.path.insert(0, str(generator_path.parent))
        try:
            spec = importlib.util.spec_from_file_location("_excel_oracle_generate_cases", generator_path)
            self.assertIsNotNone(spec)
            self.assertIsNotNone(spec.loader)
            module = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = module
            spec.loader.exec_module(module)  # type: ignore[union-attr]
        finally:
            sys.path[:] = sys_path_before
        return module

    def test_validation_rejects_volatile_functions_by_default(self) -> None:
        module = self._load_generator()

        payload = module.generate_cases()
        module._validate_against_function_catalog(payload)

        # Inject a volatile function formula; this should be rejected unless allow_volatile=True.
        bad_payload = dict(payload)
        bad_payload["cases"] = list(payload.get("cases", [])) + [
            {
                "formula": "=RAND()",
                "outputCell": "C1",
                "inputs": [],
                "tags": ["unit-test"],
                "id": "unit_test_volatile_case",
            }
        ]

        with self.assertRaises(SystemExit) as ctx:
            module._validate_against_function_catalog(bad_payload)
        msg = str(ctx.exception)
        self.assertIn("must not include volatile functions", msg)
        self.assertIn("RAND", msg)

        # Explicit override should allow volatile functions.
        module._validate_against_function_catalog(bad_payload, allow_volatile=True)


if __name__ == "__main__":
    unittest.main()

