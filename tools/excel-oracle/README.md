# Excel Oracle Compatibility Harness

This directory contains a **real Microsoft Excel** compatibility harness used to build an "oracle" dataset of expected formula results.

The intent is to continuously compare our formula engine against Excel across a growing, versioned corpus of cases.

## What this provides

- A deterministic, machine-readable **case corpus** (`tests/compatibility/excel-oracle/cases.json`)
- A Windows-only **Excel COM automation runner** that evaluates all cases in **real Excel** and exports results (`run-excel-oracle.ps1`)
- A Windows-only **function-translation extractor** to generate locale function name sources via `Range.Formula/FormulaLocal` (`extract-function-translations.ps1`)
- A Windows-only **error-literal extractor** to verify locale error spellings against real Excel (`extract-error-literals.ps1`)
- A Windows-only **structured-reference keyword probe** to inspect `[#Headers]`/`[#Data]`/etc localization via `FormulaLocal` (`extract-structured-reference-keywords.ps1`)
- A **comparison tool** that diffs engine output vs Excel output and emits a mismatch report (`compare.py`)
- A lightweight **compatibility gate** that runs the engine + comparison on a bounded subset (`compat_gate.py`)
- A GitHub Actions workflow (`.github/workflows/excel-compat.yml`) wired to run on `windows-2022` (engine validation) and optionally on a self-hosted Windows runner with Excel installed (oracle generation)

## Unified compatibility scorecard (corpus + Excel-oracle)

The Excel-oracle harness measures **calculation fidelity (L2)**. The compatibility corpus (`tools/corpus`) measures
**read (L1)** and **round-trip preservation (L4)**.

To merge both into a single markdown scorecard, run:

```bash
python tools/compat_scorecard.py --out-md compat_scorecard.md
# or:
python -m tools.compat_scorecard --out-md compat_scorecard.md
```

In CI, the repo also includes a workflow-run aggregator (`.github/workflows/compat-scorecard.yml`) that downloads the
corpus + oracle artifacts and uploads a unified `compat-scorecard` report.

The aggregator supports manual dispatch (`workflow_dispatch`) with an optional SHA input for backfills/debugging.

## Prerequisites (local generation)

To generate oracle data locally you must have:

- Windows
- Microsoft Excel installed (desktop)
- PowerShell (Windows PowerShell 5.1 or PowerShell 7+)
- Python 3 (for comparison/reporting, optional if you only generate data)

## Extract localized error-literal spellings (locale data)

Some Excel locales display error literals differently from the canonical (en-US) spellings used by
the engine (e.g. German `#WERT!` for `#VALUE!`, Spanish `#¡VALOR!` / `#¿NOMBRE?`).

The committed upstream sources live under:

`crates/formula-engine/src/locale/data/upstream/errors/*.tsv`

To extract/verify the localized spellings against a **real Excel install** for the active Excel UI
language, run (from repo root on Windows):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-error-literals.ps1 -Locale es-ES
node scripts/generate-locale-error-tsvs.mjs
```

Notes / caveats:

- Excel COM automation is Windows-only and requires Excel desktop installed.
- The output reflects the **active Excel UI language**. Install the corresponding Office language
  pack and set the Excel display language before extracting.
- The script prints the detected Excel UI locale and will warn if it does not match `-Locale`/`-LocaleId`.
- For `de-DE`/`fr-FR`/`es-ES`, the script also does a small sanity-check on a few sentinel error
  translations (e.g. `#VALUE!`) and warns if Excel appears misconfigured.
- Use `-Visible`, `-MaxErrors N`, and PowerShell's `-Verbose` switch for debugging.
- Newer error kinds (e.g. `#SPILL!`) may not exist in older Excel versions; the script will fail
  rather than emitting a misleading mapping if Excel appears not to recognize an error literal.
- After extracting/updating an upstream TSV, regenerate the committed exports:
  `node scripts/generate-locale-error-tsvs.mjs`. Runtime error-literal translation maps are loaded
  from the committed `*.errors.tsv` files (see `crates/formula-engine/src/locale/registry.rs`), and
  the Rust test `crates/formula-engine/tests/locale_error_tsv_sync.rs` enforces completeness +
  round-tripping.

## Extract localized function-name spellings (locale data)

For localized formula editing / round-tripping, the engine needs a complete mapping from canonical
function names (en-US) to the exact spelling Excel displays for a locale.

The committed translation sources live under:

`crates/formula-engine/src/locale/data/sources/*.json`

See also:

- [`crates/formula-engine/src/locale/data/README.md`](../../crates/formula-engine/src/locale/data/README.md)
  for completeness requirements, generators, and verification steps (especially for `es-ES`).

To extract a full mapping from a **real Excel install** for the active Excel UI language, run
(from repo root on Windows):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId de-DE `
  -OutPath crates/formula-engine/src/locale/data/sources/de-DE.json

# Normalize sources (omits identity mappings + enforces stable casing)
node scripts/normalize-locale-function-sources.js

# Regenerate + verify TSVs
node scripts/generate-locale-function-tsv.js
node scripts/generate-locale-function-tsv.js --check
```

Example for Spanish (`es-ES`):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId es-ES `
  -OutPath crates/formula-engine/src/locale/data/sources/es-ES.json

node scripts/normalize-locale-function-sources.js
node scripts/generate-locale-function-tsv.js
node scripts/generate-locale-function-tsv.js --check
```

Notes / caveats:

- The output reflects the **active Excel UI language**. Install the corresponding Office language
  pack and set the Excel display language before extracting.
- The script prints the detected Excel UI locale and will warn if it does not match `-LocaleId`.
- The script also does a small sanity-check on a few sentinel translations (e.g. `SUM`/`IF`) for
  `de-DE`/`es-ES`/`fr-FR` and warns if Excel appears misconfigured.
- The script warns if Excel maps multiple canonical functions to the same localized spelling (this
  would later cause `scripts/generate-locale-function-tsv.js` to fail).
- Use `-Visible`, `-MaxFunctions N`, and/or PowerShell’s `-Verbose` switch for debugging.
  - Note: `-MaxFunctions` is for debugging only; do not commit partial sources.
- `sources/<locale>.json` is expected to come from this extractor whenever possible. Avoid replacing
  `sources/es-ES.json` with partial online translation tables; missing entries silently fall back to
  English in the generated TSVs.
- For `es-ES`, treat a “complete” extraction as the extractor writing **one translation per canonical
  function** (before normalization) and not reporting a large skipped set.
  - If the script reports skipped functions, those functions will fall back to English in the
    generated TSVs; use a newer Excel build / correct language pack rather than committing a partial
    mapping.
- Before committing, normalize the extracted JSON sources to omit identity mappings and ensure
  deterministic casing:
  - `node scripts/normalize-locale-function-sources.js` (or `pnpm normalize:locale-function-sources`)
  - CI-style check: `node scripts/normalize-locale-function-sources.js --check` (or `pnpm check:locale-function-sources`)
  - Note: after normalization, the committed `sources/<locale>.json` will typically contain fewer
    entries than `shared/functionCatalog.json`, since identity mappings are omitted.
- After extracting, regenerate + verify with:
  - `node scripts/generate-locale-function-tsv.js`
  - `node scripts/generate-locale-function-tsv.js --check`
  - and spot-check that Spanish financial functions like `NPV`/`IRR` localize (e.g. `VNA`/`TIR`).
- Optional (recommended when touching locale data): run the Rust guard-rail tests:
  - `bash scripts/cargo_agent.sh test -p formula-engine --test locale_function_tsv_completeness`
  - `bash scripts/cargo_agent.sh test -p formula-engine --test locale_es_es_function_sentinels`

Verification checklist note (especially `es-ES`):

- Do **not** populate `es-ES` from third-party lists/websites — they are often incomplete/outdated.
  Always extract from a real Excel install with `extract-function-translations.ps1`.
- After regenerating `crates/formula-engine/src/locale/data/<locale>.tsv`, spot-check sentinel
  translations (e.g. Spanish `SUM → SUMA`, `IF → SI`, `NPV → VNA`, `IRR → TIR`, …) and sanity-check
  that the TSV does not contain a suspiciously large number of identity mappings.
- Ensure there are no localized-name collisions (the extractor warns; the TSV generator fails).

See `crates/formula-engine/src/locale/data/README.md` for the full locale TSV contract and the
verification checklist.

## CI note (Excel availability)

GitHub-hosted Windows runners (for example `windows-2022`) typically **do not include Microsoft Excel**. To generate oracle data in CI you generally need a **self-hosted Windows runner** with Excel installed.

If you commit a pinned oracle dataset (see below), CI can still validate the engine even when Excel is not available.

See `tools/excel-oracle/self-hosted-runner.md` for notes on running Excel COM automation in CI.

## Case corpus

The canonical case list lives at:

`tests/compatibility/excel-oracle/cases.json`

It is generated by:

```powershell
python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json
```

The generator is deterministic; committing `cases.json` makes CI stable and reviewable.

For local debugging, the generator also supports:

```powershell
python tools/excel-oracle/generate_cases.py --include-volatile --out /tmp/cases.json
```

This opt-in mode includes volatile functions like `CELL`/`INFO` (which cannot be pinned/stably
compared). The generated cases file is still deterministic, but the results are not stable/pinnable,
so it must not be committed or used for pinned datasets.
Currently this flag only enables the `CELL`/`INFO` volatile debug cases; other volatile functions
(e.g. `RAND`, `NOW`) remain forbidden.

As an additional safety guard, `generate_cases.py` refuses to overwrite the committed
`tests/compatibility/excel-oracle/cases.json` when `--include-volatile` is set.

The generator also validates the corpus against `shared/functionCatalog.json` to ensure:

- every `non_volatile` catalog function is exercised by at least one **case formula** (`case.formula`)
- `volatile` catalog functions are excluded (so pinned comparisons remain deterministic)

The Rust test suite enforces the same invariants (see
`crates/formula-engine/tests/excel_oracle_coverage.rs`) so drift is caught even if
`cases.json` is edited without re-running the generator.

## How to add oracle coverage for a new function

1) Add at least one new case to the appropriate module under:

`tools/excel-oracle/case_generators/`

Common buckets:

- `arith.py` (operators)
- `math.py`
- `engineering.py`
- `statistical.py` (aggregates + criteria semantics + stats/regression)
- `logical.py`
- `coercion.py` (type/boolean/blank coercion semantics)
- `text.py`
- `date_time.py`
- `lookup.py`
- `database.py`
- `financial.py`
- `spill.py` (dynamic arrays / spill behavior)
- `info.py`
- `lambda_cases.py` (LAMBDA / LET / MAP, etc.)
- `errors.py` (error creation/propagation)

Each module exposes `generate(cases, *, add_case, CellInput, ...) -> None` and is invoked
in a deterministic order from `tools/excel-oracle/generate_cases.py`.

Guidelines:

- Keep the corpus **deterministic**: do not add volatile functions (e.g. `RAND`, `NOW`).
- Keep the corpus **small**: the generator hard-caps at 2000 cases so it can run in real Excel in CI.
- Prefer locale-independent formulas/inputs where possible (avoid ambiguous date strings).

2) Regenerate the committed case corpus:

```bash
python tools/excel-oracle/generate_cases.py --out tests/compatibility/excel-oracle/cases.json
```

3) Validate:

```bash
python -m unittest tools/excel-oracle/tests/test_*.py
```

If you intentionally changed the case corpus, you may also need to re-pin the Excel dataset
(`tools/excel-oracle/pin_dataset.py`) so `tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json`
stays in sync.

## Generate oracle dataset from Excel

From repo root on Windows:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

Note: the generated dataset includes `caseSet.path` metadata. If you pass an absolute `-CasesPath`,
the script normalizes it to a portable, privacy-safe relative path (when it can detect a repo-relative
suffix like `tests/...` or `tools/...`).

Tip: pass `-DryRun` to see how many cases would be selected by the tag filters / `-MaxCases` without starting Excel.

To generate only a subset of cases (by tag):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath tests/compatibility/excel-oracle/datasets/excel-oracle.json `
  -IncludeTags SUM,IF,cmp `
  -ExcludeTags spill,dynarr
```

Note: `-IncludeTags` uses **OR semantics** (a case is included if it contains *any* of the included
tags). To require that a case contains *all* tags (AND semantics), use `-RequireTags`:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tests/compatibility/excel-oracle/cases.json `
  -OutPath tests/compatibility/excel-oracle/datasets/excel-oracle.json `
  -RequireTags odd_coupon,basis4
```

To generate only the **long odd-coupon** stub scenarios (`ODDF*` / `ODDL*`) for quick iteration / pinning,
use the small subset corpus:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tools/excel-oracle/odd_coupon_long_stub_cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

To generate only the **odd-coupon negative yield / negative coupon** validation scenarios, use:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tools/excel-oracle/odd_coupon_validation_cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

This subset corresponds to the cases tagged `odd_coupon_validation` in the canonical corpus.

To generate only the **odd-coupon boundary** date-equality scenarios (e.g. `issue == settlement`),
use:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tools/excel-oracle/odd_coupon_boundary_cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

To generate only the **odd-coupon invalid schedule** scenarios (cases covering schedule alignment /
misalignment; some return `#NUM!`, while others are accepted by the engine today and should be
validated against real Excel), use:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tools/excel-oracle/odd_coupon_invalid_schedule_cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

This subset corresponds to the cases tagged `odd_coupon` + `invalid_schedule` in the canonical corpus.

To generate only the **odd-coupon basis=4** scenarios (European 30/360), use:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/run-excel-oracle.ps1 `
  -CasesPath tools/excel-oracle/odd_coupon_basis4_cases.json `
  -OutPath  tests/compatibility/excel-oracle/datasets/excel-oracle.json
```

This subset corresponds to the cases tagged `odd_coupon` + `basis4` in the canonical corpus.

To regenerate the derived odd-coupon subset corpora (boundary + validation + long-stub + basis4 + invalid-schedule) from the
canonical corpus, run:

```bash
python tools/excel-oracle/regenerate_subset_corpora.py
```

To only verify the subset corpora are up to date (without rewriting files), run:

```bash
python tools/excel-oracle/regenerate_subset_corpora.py --check
```

To preview what would be written (case counts + paths) without writing files, run:

```bash
python tools/excel-oracle/regenerate_subset_corpora.py --dry-run
```

Note: this subset corpus reuses the **canonical case IDs** from `tests/compatibility/excel-oracle/cases.json`,
so you can map results back to the full corpus by `caseId` (useful when updating/pinning datasets).

The output JSON includes:

- Excel version/build metadata (because behavior can differ between Excel versions)
- A SHA-256 of the case corpus used
- Results encoded with a stable, typed representation (see below)

## Pin an oracle dataset (optional, for CI without Excel)

If you want CI to validate without running Excel, you can commit a pinned dataset file:

```bash
python tools/excel-oracle/pin_dataset.py \
  --dataset tests/compatibility/excel-oracle/datasets/excel-oracle.json \
  --pinned tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --versioned-dir tests/compatibility/excel-oracle/datasets/versioned
```

To validate the dataset and preview the output paths without writing files, use:

```bash
python tools/excel-oracle/pin_dataset.py --dataset /path/to/dataset.json --dry-run
```

Note: `pin_dataset.py` enforces that the dataset includes Excel version/build/OS metadata
from COM automation when pinning a real Excel dataset, to avoid accidentally pinning
something that simply sets `source.kind="excel"`.

`pin_dataset.py` also supports pinning a **synthetic CI baseline** from the in-repo Rust
engine (`crates/formula-excel-oracle`, where `source.kind == "formula-engine"`). In that
mode it re-tags the dataset as `source.kind="excel"` with `"unknown"` metadata and
embeds the original engine metadata under `source.syntheticSource`.

The workflow prefers `excel-oracle.pinned.json` if present.

To force Excel generation in the workflow when a pinned dataset exists, run the workflow manually (`workflow_dispatch`) and set `oracle_source=generate`.

## Regenerate the synthetic CI baseline (no Excel required)

When adding new deterministic functions (e.g. new STAT functions), you often need to update
multiple committed artifacts together to keep CI green:

- `shared/functionCatalog.json` (+ `shared/functionCatalog.mjs`)
- `tests/compatibility/excel-oracle/cases.json`
- `tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json`

To regenerate all of them from the current `formula-engine` implementation:

```bash
python tools/excel-oracle/regenerate_synthetic_baseline.py
```

To preview the commands it would run (without executing them or writing files), use:

```bash
python tools/excel-oracle/regenerate_synthetic_baseline.py --dry-run
```

### Incremental pinned dataset updates (merge-friendly)

When you *only add new cases* to `cases.json` (i.e. existing case IDs remain valid), regenerating the
entire pinned dataset can create very large diffs and frequent merge conflicts.

You can instead update the pinned dataset incrementally to fill in only the missing case results:

```bash
python tools/excel-oracle/update_pinned_dataset.py
```

To preview what would happen (missing cases, whether the engine would run, etc.) without writing
files or running Cargo, use:

```bash
python tools/excel-oracle/update_pinned_dataset.py --dry-run
```

By default this also refreshes the matching **versioned** dataset copy under:

* `tests/compatibility/excel-oracle/datasets/versioned/`

so `tools/excel-oracle/compat_gate.py` (which prefers the versioned dataset when present) stays in
sync. To update only the pinned file, pass `--no-versioned`.

If you generated **real Excel** results for a subset of cases and want to overwrite the synthetic
baseline values in the pinned dataset (while keeping the rest of the corpus unchanged), pass
`--merge-results` and `--overwrite-existing`:

```bash
python tools/excel-oracle/update_pinned_dataset.py \
  --merge-results /path/to/excel-results.json \
  --overwrite-existing \
  --no-engine
```

For a one-command flow that runs Excel on a subset corpus and patches the pinned dataset, see:

* `tools/excel-oracle/patch-pinned-dataset-with-excel.ps1`

Tip: pass `-DryRun` to preview the commands without running Excel or modifying the pinned dataset.

Example (patch only the odd-coupon negative yield / negative coupon validation scenarios):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_validation_cases.json
```

Example (patch only the odd-coupon boundary date-equality scenarios):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_boundary_cases.json
```

Example (patch only the odd-coupon basis=4 (European 30/360) scenarios):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_basis4_cases.json
```

Example (patch only the odd-coupon long-stub scenarios):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_long_stub_cases.json
```

Example (patch only the odd-coupon schedule alignment / misalignment scenarios (invalid schedule cases)):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_invalid_schedule_cases.json
```

You can also patch by **tag filter** without a dedicated subset file by running against the
canonical corpus and passing through `-IncludeTags` (OR semantics), `-RequireTags` (AND semantics),
and/or `-ExcludeTags`:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tests/compatibility/excel-oracle/cases.json `
  -IncludeTags odd_coupon_validation
```

Example (patch only the odd-coupon `basis=4` cases using AND tag filtering):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tests/compatibility/excel-oracle/cases.json `
  -RequireTags odd_coupon,basis4
```

This script patches the pinned dataset by invoking `update_pinned_dataset.py`, which preserves
existing results and updates the `caseSet.sha256`/`caseSet.count` metadata. The patch flow uses
`--no-engine`, so only the provided Excel results are merged/overwritten (no engine fallback).

When you patch a synthetic baseline with **real Excel** results, `update_pinned_dataset.py` records
a small provenance entry under `source.patches` in the pinned dataset so it’s clear which case IDs
were overwritten (and which Excel version/build produced them).

Important: by default the updater refuses to fill missing cases if the pinned dataset appears to be
generated by real Excel (no `source.syntheticSource` metadata), because that would mix engine results
into an Excel oracle dataset. In that scenario, generate new Excel results and merge them with
`--merge-results` (or use `--force-engine` if you explicitly want a synthetic baseline).

This will:

1. Regenerate the function catalog from `formula-engine`'s inventory registry.
2. Regenerate the oracle case corpus (and validate coverage).
3. Evaluate the full corpus using `crates/formula-excel-oracle`.
4. Pin the results as a synthetic dataset for CI.

## Compare formula-engine output vs Excel oracle

### One-command gate (CI-friendly)

From repo root:

```bash
python tools/excel-oracle/compat_gate.py
```

For defense in depth when generating reports in less-trusted environments, you can enable privacy mode:

```bash
python tools/excel-oracle/compat_gate.py --privacy-mode private
```

This hashes absolute filesystem paths embedded in the mismatch report (for example Windows paths with
usernames). Repo-relative paths remain readable.

The gate supports tier presets:

```bash
python tools/excel-oracle/compat_gate.py --tier smoke   # default, CI-friendly slice
python tools/excel-oracle/compat_gate.py --tier p0      # broader common-function slice
python tools/excel-oracle/compat_gate.py --tier full    # full corpus (no include-tag filtering)
```

To inspect which datasets/tags will be selected (and to see the exact engine + compare commands)
without running `cargo`, use:

```bash
python tools/excel-oracle/compat_gate.py --dry-run --tier smoke --max-cases 10
```

See `tests/compatibility/excel-oracle/README.md` for the exact tag presets and
recommended runtime tradeoffs.

This runs the in-repo engine adapter (`crates/formula-excel-oracle`) against a curated tag set,
compares against the pinned dataset in `tests/compatibility/excel-oracle/datasets/versioned/`,
writes reports under `tests/compatibility/excel-oracle/reports/`, and exits non-zero on mismatch.

Note on **synthetic baselines**: the default pinned dataset in this repo may be a *synthetic CI baseline*
(`source.syntheticSource`), meaning it is **not** generated by real Microsoft Excel. `compat_gate.py`
prints an explicit warning when it detects a synthetic expected dataset. To enforce that CI is using a
real Excel dataset (for example on a self-hosted Windows runner that can run `run-excel-oracle.ps1`),
pass:

```bash
python tools/excel-oracle/compat_gate.py --require-real-excel
```

### Manual flow

1) Produce engine results JSON (same schema as Excel output). The intended flow is that your engine exposes a CLI that can evaluate the case corpus and emit results.

This repo includes a reference implementation for the Rust `formula-engine` at:

`crates/formula-excel-oracle/` (run with `cargo run -p formula-excel-oracle -- --cases ... --out ...`).

It also supports `--include-tag` / `--exclude-tag` for evaluating a filtered subset of the corpus.

2) Compare:

```bash
python tools/excel-oracle/compare.py \
  --cases tests/compatibility/excel-oracle/cases.json \
  --expected tests/compatibility/excel-oracle/datasets/excel-oracle.json \
  --actual tests/compatibility/excel-oracle/datasets/engine-results.json \
  --report tests/compatibility/excel-oracle/reports/mismatch-report.json
```

If you are generating reports on a machine where paths may include sensitive usernames/mount points,
use:

```bash
python tools/excel-oracle/compare.py --privacy-mode private ...
```

In `privacy-mode=private`, absolute filesystem paths in the report metadata are hashed; relative paths
are preserved.

To preview how many cases would be compared (after tag filtering / `--max-cases`) without writing a report file, use:

```bash
python tools/excel-oracle/compare.py --dry-run \
  --cases tests/compatibility/excel-oracle/cases.json \
  --expected tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --actual tests/compatibility/excel-oracle/datasets/engine-results.json \
  --report tests/compatibility/excel-oracle/reports/mismatch-report.json
```

The report includes `caseId`, `formula`, `inputs`, `expected`, `actual`, and a reason.

`compare.py` also verifies that the `caseSet.sha256` embedded in the datasets matches the current `cases.json`, to prevent stale-oracle comparisons.

### Tag filtering

Cases include `tags` (e.g. `["logical","IF"]`). You can restrict comparisons to a subset:

```bash
python tools/excel-oracle/compare.py \
  --cases tests/compatibility/excel-oracle/cases.json \
  --expected tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --actual tests/compatibility/excel-oracle/datasets/engine-results.json \
  --report tests/compatibility/excel-oracle/reports/mismatch-report.json \
  --include-tag IF --include-tag SUM --include-tag cmp
```

### Numeric tolerances (iterative functions)

`compare.py` defaults to tight numeric tolerances (`abs=rel=1e-9`). Some functions are inherently
iterative (for example yield solvers), and can differ from Excel by small floating point amounts even
when the math is correct.

You can override numeric tolerances for tagged subsets without loosening the entire corpus:

```bash
python tools/excel-oracle/compare.py \
  --cases    tests/compatibility/excel-oracle/cases.json \
  --expected tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --actual   tests/compatibility/excel-oracle/datasets/engine-results.json \
  --report   tests/compatibility/excel-oracle/reports/mismatch-report.json \
  --tag-abs-tol odd_coupon=1e-6 \
  --tag-rel-tol odd_coupon=1e-6
```

Note: `tools/excel-oracle/compat_gate.py` already applies `odd_coupon=1e-6` by default.

## Value encoding

Excel values are encoded to avoid ambiguity between:

- blank vs 0
- numbers vs error codes
- scalar vs spilled array results

See `tools/excel-oracle/value-encoding.md`.

## Schemas

JSON Schemas (for editor validation) live in `tools/excel-oracle/schemas/`.
