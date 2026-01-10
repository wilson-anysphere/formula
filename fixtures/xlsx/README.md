# XLSX Fixture Corpus

These files are used by the round-trip validation harness (`crates/xlsx-diff`).

Goals:
- Keep fixtures **small** and **in-repo** so CI is fast and deterministic.
- Cover representative XLSX features incrementally (basic cells, formulas, styles, etc.).

## Layout

```
fixtures/xlsx/
  basic/
  formulas/
  styles/
  conditional-formatting/
  charts/
  pivots/
  macros/
```

`charts/`, `pivots/`, and `macros/` are currently scaffolds for future expansion.

## Regenerating the initial fixtures

The initial `.xlsx` files are generated without external dependencies:

```bash
python3 fixtures/xlsx/generate_fixtures.py
```

The generator uses deterministic timestamps so diffs are stable.

