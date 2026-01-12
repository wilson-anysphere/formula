from __future__ import annotations

import hashlib
import importlib.util
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest import mock


class CompatGateDatasetSelectionTests(unittest.TestCase):
    def _load_compat_gate(self):
        compat_gate_py = Path(__file__).resolve().parents[1] / "compat_gate.py"
        self.assertTrue(compat_gate_py.is_file(), f"compat_gate.py not found at {compat_gate_py}")

        spec = importlib.util.spec_from_file_location("excel_oracle_compat_gate", compat_gate_py)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_default_expected_dataset_prefers_versioned_match_by_cases_sha(self) -> None:
        # Ensure compat_gate selects the expected dataset by matching the cases.json hash
        # (via the versioned filename suffix), not by lexicographic ordering.
        compat_gate = self._load_compat_gate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            old_cwd = Path.cwd()
            try:
                os.chdir(tmp_path)

                cases_path = tmp_path / "cases.json"
                cases_path.write_text('{"schemaVersion":1,"cases":[]}\n', encoding="utf-8", newline="\n")
                sha8 = hashlib.sha256(cases_path.read_bytes()).hexdigest()[:8]

                versioned_dir = tmp_path / "tests/compatibility/excel-oracle/datasets/versioned"
                versioned_dir.mkdir(parents=True, exist_ok=True)

                older = versioned_dir / f"excel-16.0-build-1-cases-{sha8}.json"
                newer = versioned_dir / f"excel-16.0-build-2-cases-{sha8}.json"
                older.write_text("{}", encoding="utf-8", newline="\n")
                newer.write_text("{}", encoding="utf-8", newline="\n")

                selected = compat_gate._default_expected_dataset(cases_path=cases_path)
                self.assertEqual(selected.resolve(), newer.resolve())
            finally:
                os.chdir(old_cwd)

    def test_default_expected_dataset_prefers_real_excel_over_unknown(self) -> None:
        compat_gate = self._load_compat_gate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            old_cwd = Path.cwd()
            try:
                os.chdir(tmp_path)

                cases_path = tmp_path / "cases.json"
                cases_path.write_text('{"schemaVersion":1,"cases":[]}\n', encoding="utf-8", newline="\n")
                sha8 = hashlib.sha256(cases_path.read_bytes()).hexdigest()[:8]

                versioned_dir = tmp_path / "tests/compatibility/excel-oracle/datasets/versioned"
                versioned_dir.mkdir(parents=True, exist_ok=True)

                unknown = versioned_dir / f"excel-unknown-build-unknown-cases-{sha8}.json"
                real = versioned_dir / f"excel-16.0-build-2-cases-{sha8}.json"
                unknown.write_text("{}", encoding="utf-8", newline="\n")
                real.write_text("{}", encoding="utf-8", newline="\n")

                selected = compat_gate._default_expected_dataset(cases_path=cases_path)
                self.assertEqual(selected.resolve(), real.resolve())
            finally:
                os.chdir(old_cwd)

    def test_default_expected_dataset_falls_back_to_pinned(self) -> None:
        compat_gate = self._load_compat_gate()

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            old_cwd = Path.cwd()
            try:
                os.chdir(tmp_path)

                cases_path = tmp_path / "cases.json"
                cases_path.write_text('{"schemaVersion":1,"cases":[]}\n', encoding="utf-8", newline="\n")

                pinned = tmp_path / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
                pinned.parent.mkdir(parents=True, exist_ok=True)
                pinned.write_text("{}", encoding="utf-8", newline="\n")

                selected = compat_gate._default_expected_dataset(cases_path=cases_path)
                self.assertEqual(selected.resolve(), pinned.resolve())
            finally:
                os.chdir(old_cwd)

    def test_repo_pinned_dataset_matches_selected_unknown_versioned_dataset(self) -> None:
        """Ensure the repo's pinned + versioned datasets don't drift.

        `compat_gate.py` prefers a versioned dataset if it exists for the current `cases.json` hash.
        When that dataset is an `unknown-build-unknown` synthetic baseline, the versioned file should
        be an exact copy of `excel-oracle.pinned.json` (as produced by `pin_dataset.py`).
        """

        compat_gate = self._load_compat_gate()

        repo_root = Path(__file__).resolve().parents[3]
        cases_path = repo_root / "tests/compatibility/excel-oracle/cases.json"
        pinned_path = (
            repo_root / "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
        )
        self.assertTrue(cases_path.is_file(), f"cases.json not found at {cases_path}")
        self.assertTrue(pinned_path.is_file(), f"pinned dataset not found at {pinned_path}")

        def _sha256_file(path: Path) -> str:
            h = hashlib.sha256()
            with path.open("rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
            return h.hexdigest()

        old_cwd = Path.cwd()
        try:
            os.chdir(repo_root)
            selected = compat_gate._default_expected_dataset(cases_path=cases_path).resolve()
            self.assertTrue(selected.is_file(), f"selected dataset does not exist: {selected}")

            if "-unknown-build-unknown-" in selected.name:
                self.assertEqual(
                    _sha256_file(selected),
                    _sha256_file(pinned_path),
                    "selected versioned dataset should match excel-oracle.pinned.json "
                    "(run tools/excel-oracle/pin_dataset.py with --versioned-dir to regenerate)",
                )
        finally:
            os.chdir(old_cwd)


class CompatGateTierPresetTests(unittest.TestCase):
    def _load_compat_gate(self):
        compat_gate_py = Path(__file__).resolve().parents[1] / "compat_gate.py"
        self.assertTrue(compat_gate_py.is_file(), f"compat_gate.py not found at {compat_gate_py}")

        spec = importlib.util.spec_from_file_location("excel_oracle_compat_gate", compat_gate_py)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_full_tier_defaults_to_no_include_tags(self) -> None:
        compat_gate = self._load_compat_gate()

        include_tags = compat_gate._effective_include_tags(tier="full", user_include_tags=[])
        self.assertEqual(include_tags, [])

        cmd = compat_gate._build_engine_cmd(
            cases_path=Path("cases.json"),
            actual_path=Path("engine-results.json"),
            max_cases=0,
            include_tags=include_tags,
            exclude_tags=["error"],
            use_cargo_agent=False,
        )
        self.assertNotIn("--include-tag", cmd)
        self.assertIn("--exclude-tag", cmd)
        self.assertIn("error", cmd)

    def test_smoke_and_p0_tiers_include_thai_tag(self) -> None:
        compat_gate = self._load_compat_gate()

        smoke_tags = compat_gate._effective_include_tags(tier="smoke", user_include_tags=[])
        p0_tags = compat_gate._effective_include_tags(tier="p0", user_include_tags=[])
        self.assertIn("thai", smoke_tags)
        self.assertIn("thai", p0_tags)

    def test_smoke_tier_includes_coupon_schedule_tag(self) -> None:
        compat_gate = self._load_compat_gate()

        smoke_tags = compat_gate._effective_include_tags(tier="smoke", user_include_tags=[])
        self.assertIn("coupon_schedule", smoke_tags)

    def test_user_include_tag_overrides_tier_presets(self) -> None:
        compat_gate = self._load_compat_gate()

        include_tags = compat_gate._effective_include_tags(tier="smoke", user_include_tags=[" add ", "", "cmp"])
        self.assertEqual(include_tags, ["add", "cmp"])

        cmd = compat_gate._build_compare_cmd(
            cases_path=Path("cases.json"),
            expected_path=Path("expected.json"),
            actual_path=Path("actual.json"),
            report_path=Path("report.json"),
            max_cases=0,
            include_tags=include_tags,
            exclude_tags=[],
            max_mismatch_rate=0.0,
            abs_tol=1e-9,
            rel_tol=1e-9,
            tag_abs_tol=[],
            tag_rel_tol=[],
        )
        # Ensure we only pass through the explicit tags.
        self.assertEqual(cmd.count("--include-tag"), 2)
        self.assertIn("add", cmd)
        self.assertIn("cmp", cmd)


class CompatGateDryRunTests(unittest.TestCase):
    def _load_compat_gate(self):
        compat_gate_py = Path(__file__).resolve().parents[1] / "compat_gate.py"
        self.assertTrue(compat_gate_py.is_file(), f"compat_gate.py not found at {compat_gate_py}")

        spec = importlib.util.spec_from_file_location("excel_oracle_compat_gate_dry_run", compat_gate_py)
        assert spec is not None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        assert spec.loader is not None
        spec.loader.exec_module(module)
        return module

    def test_dry_run_does_not_invoke_subprocesses(self) -> None:
        compat_gate = self._load_compat_gate()
        stdout = StringIO()

        repo_root = Path(__file__).resolve().parents[3]
        old_cwd = Path.cwd()
        old_argv = sys.argv[:]
        try:
            os.chdir(repo_root)
            sys.argv = [
                "compat_gate.py",
                "--tier",
                "smoke",
                "--max-cases",
                "1",
                "--dry-run",
            ]
            with mock.patch.object(compat_gate.subprocess, "run") as run_mock, redirect_stdout(stdout):
                rc = compat_gate.main()
            run_mock.assert_not_called()
            self.assertEqual(rc, 0)
        finally:
            sys.argv = old_argv
            os.chdir(old_cwd)

        out = stdout.getvalue()
        self.assertIn("engine_cmd:", out)
        self.assertIn("compare_cmd:", out)


if __name__ == "__main__":
    unittest.main()
