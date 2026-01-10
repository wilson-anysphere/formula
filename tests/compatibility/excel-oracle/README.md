# Excel Oracle compatibility tests

This directory holds the **inputs** to the Excel-oracle compatibility harness.

## Files

- `cases.json` — curated (~1k) formula + input-grid cases.
- `datasets/` — output datasets (Excel oracle and engine results). These are typically generated (and may be uploaded as CI artifacts).
- `reports/` — mismatch reports produced by `tools/excel-oracle/compare.py`.

## Regenerating cases

From repo root:

```bash
python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json
```

