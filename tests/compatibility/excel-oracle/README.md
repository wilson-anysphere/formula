# Excel Oracle compatibility tests

This directory holds the **inputs** to the Excel-oracle compatibility harness.

## Files

- `cases.json` — curated (~1k) formula + input-grid cases.
- `datasets/` — output datasets (Excel oracle and engine results). These are typically generated (and may be uploaded as CI artifacts).
- `datasets/excel-oracle.pinned.json` — optional pinned Excel oracle dataset (commit this if you want CI to validate without running Excel).
- `datasets/versioned/` — optional version-tagged pinned datasets (useful when Excel behavior differs across versions/builds).
- `reports/` — mismatch reports produced by `tools/excel-oracle/compare.py`.

## Tags and filtering

Each case has `tags` that can be used to include/exclude subsets when comparing.
This is useful for incrementally expanding coverage (e.g. exclude `spill` until
dynamic arrays are implemented in the engine).

## Regenerating cases

From repo root:

```bash
python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json
```
