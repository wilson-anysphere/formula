# Excel Oracle compatibility tests

This directory contains the **case corpus** and **pinned Excel results** used to
continuously validate `crates/formula-engine` against real Excel behavior.

## Quick start (one command)

From repo root:

```bash
python tools/excel-oracle/compat_gate.py
```

This runs:

1. `cargo run -p formula-excel-oracle` → writes `datasets/engine-results.json`
2. `python tools/excel-oracle/compare.py` → writes `reports/mismatch-report.json`

It also writes a human-readable summary to:

* `reports/summary.md`

and exits non-zero if mismatches exceed the configured threshold.

## Repository layout

- `cases.json` — curated (~1.4k) formula + input-grid cases (deterministic).
- `datasets/` — results datasets:
  - `excel-oracle.pinned.json` — pinned Excel dataset for CI (no Excel needed).
  - `versioned/` — version-tagged pinned datasets (useful when Excel behavior differs across versions/builds).
  - `engine-results.json` — generated locally/CI by the engine runner.
- `reports/` — mismatch reports produced by `compare.py`.

Note: the repo may include a **small pinned dataset** to keep CI fast. For the
true Excel oracle, regenerate the dataset with the Windows + Excel runner and
pin it (see below).

The case corpus generator (`tools/excel-oracle/generate_cases.py`) validates that:

- every `non_volatile` function in `shared/functionCatalog.json` has at least one case
- `volatile` catalog functions are excluded (keeps pinned oracle comparisons deterministic)

## Tags and filtering

Each case has `tags` that can be used to include/exclude subsets when evaluating
or comparing.

This lets CI stay fast (start with a small tag set) while still enabling full
corpus runs locally or on a self-hosted Windows runner.

`compat_gate.py` defaults to a curated tag set (basic arithmetic/comparison + a
few baseline functions), plus a small amount of **dynamic array spill coverage**
(`range`, `TRANSPOSE`, `SEQUENCE`).

The default tag slice also includes representative **value coercion / conversion**
coverage (tagged `coercion` / `VALUE` / `DATEVALUE` / `TIMEVALUE`), so changes to
text→number/date/time semantics are exercised in CI even before a full Excel
oracle dataset is generated.

## Regenerate the case corpus

From repo root:

```bash
python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json
```

## Generate Excel oracle results (Windows + Excel required)

On a Windows machine with Microsoft Excel installed:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

Then pin the dataset for CI (and optionally write a version-tagged copy):

```bash
python tools/excel-oracle/pin_dataset.py \
  --dataset tests/compatibility/excel-oracle/datasets/excel-oracle.json \
  --pinned tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --versioned-dir tests/compatibility/excel-oracle/datasets/versioned
```

## Run the engine runner directly

```bash
cargo run -p formula-excel-oracle -- \
  --cases tests/compatibility/excel-oracle/cases.json \
  --out   tests/compatibility/excel-oracle/datasets/engine-results.json \
  --include-tag add --include-tag sub --include-tag mul --include-tag div
```

## Run comparison directly

```bash
python tools/excel-oracle/compare.py \
  --cases    tests/compatibility/excel-oracle/cases.json \
  --expected tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --actual   tests/compatibility/excel-oracle/datasets/engine-results.json \
  --report   tests/compatibility/excel-oracle/reports/mismatch-report.json
```
