from __future__ import annotations

import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class GenerateCasesVolatileFlagTests(unittest.TestCase):
    def _load_generator_module(self, *, repo_root: Path):
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

    def test_include_volatile_generates_deterministic_superset(self) -> None:
        repo_root = Path(__file__).resolve().parents[3]
        gen_py = repo_root / "tools/excel-oracle/generate_cases.py"
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"

        self.assertTrue(gen_py.is_file(), f"generate_cases.py not found at {gen_py}")
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")

        committed = json.loads(cases_path.read_text(encoding="utf-8"))
        committed_ids = {
            c.get("id") for c in committed.get("cases", []) if isinstance(c, dict) and isinstance(c.get("id"), str)
        }

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            out1 = tmp / "cases-volatile-1.json"
            out2 = tmp / "cases-volatile-2.json"

            for out in (out1, out2):
                proc = subprocess.run(
                    [sys.executable, str(gen_py), "--include-volatile", "--out", str(out)],
                    capture_output=True,
                    text=True,
                    cwd=str(repo_root),
                )
                if proc.returncode != 0:
                    self.fail(
                        f"generate_cases.py --include-volatile exited {proc.returncode}\n"
                        f"stdout:\n{proc.stdout}\n"
                        f"stderr:\n{proc.stderr}"
                    )

            # Determinism: volatile mode should still be byte-stable (even if the formulas themselves
            # are volatile at runtime).
            self.assertEqual(
                hashlib.sha256(out1.read_bytes()).hexdigest(),
                hashlib.sha256(out2.read_bytes()).hexdigest(),
                "--include-volatile corpus output should be deterministic",
            )

            volatile_payload = json.loads(out1.read_text(encoding="utf-8"))
            volatile_cases = volatile_payload.get("cases", [])
            self.assertIsInstance(volatile_cases, list)

            volatile_ids = {
                c.get("id") for c in volatile_cases if isinstance(c, dict) and isinstance(c.get("id"), str)
            }

            # Volatile corpus should be a strict superset of the committed deterministic corpus.
            self.assertTrue(committed_ids.issubset(volatile_ids))
            extra_ids = volatile_ids.difference(committed_ids)
            self.assertGreater(len(extra_ids), 0, "Expected --include-volatile to add cases")

            # Ensure we actually emitted INFO/CELL formulas (the main motivation for the flag).
            formulas = [c.get("formula", "") for c in volatile_cases if isinstance(c, dict) and isinstance(c.get("formula"), str)]
            self.assertTrue(any("=CELL(" in f for f in formulas), "Expected at least one CELL() case in --include-volatile mode")
            self.assertTrue(any("=INFO(" in f for f in formulas), "Expected at least one INFO() case in --include-volatile mode")

            # The only extras should be the volatile cases, so keep this invariant loose but meaningful.
            self.assertTrue(
                all(cid.startswith(("cell_", "info_")) for cid in extra_ids),
                f"Unexpected extra case IDs in --include-volatile mode: {sorted(list(extra_ids))[:10]}",
            )

            # Ensure the corpus does not accidentally grow additional volatile functions over time
            # without an explicit opt-in.
            module = self._load_generator_module(repo_root=repo_root)
            used_any: set[str] = set()
            for case in volatile_cases:
                if not isinstance(case, dict):
                    continue
                used_any.update(module._extract_function_names(case.get("formula")))
                inputs = case.get("inputs", [])
                if isinstance(inputs, list):
                    for cell_input in inputs:
                        if isinstance(cell_input, dict):
                            used_any.update(module._extract_function_names(cell_input.get("formula")))

            catalog = json.loads((repo_root / "shared" / "functionCatalog.json").read_text(encoding="utf-8"))
            volatile = {fn.get("name", "").upper() for fn in catalog.get("functions", []) if isinstance(fn, dict) and fn.get("volatility") == "volatile"}
            present_volatile = sorted(used_any.intersection(volatile))
            self.assertEqual(
                set(present_volatile),
                {"CELL", "INFO"},
                f"Unexpected volatile functions present in --include-volatile corpus: {present_volatile}",
            )


if __name__ == "__main__":
    unittest.main()
