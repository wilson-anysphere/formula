# Compatibility corpus ("real-world XLSX/XLSB zoo")

## Why this exists

Excel compatibility is won and lost on **real workbooks**. The goal of the compatibility corpus is to:

- collect problematic/representative `.xlsx`/`.xlsm`/`.xlsb` files (from users and internal sources)
- sanitize them with **privacy controls**
- run automated triage (open → recalc → render-smoke → round-trip save → diff)
- surface regressions quickly in CI, with scorecards that can be tracked over time

This is the foundation for a compatibility dashboard and an "XLSX zoo" that grows with the product.

---

## Corpus layout

This repo supports two corpora:

### 1) Public subset (runs in PR CI)

`tools/corpus/public/`

- Contains only **non-sensitive** fixtures.
- Small files may be stored as `*.xlsx.b64` (base64 text) to avoid committing binaries.
- Small XLSB fixtures may be stored as `*.xlsb.b64` (base64 text) as well.
- Expected pass/fail states are tracked in `tools/corpus/public/expectations.json`.

### 2) Private corpus (runs in scheduled CI)

`tools/corpus/private/` *(gitignored)*

Populated locally via `tools/corpus/ingest.py`, or in CI by downloading a private archive.

Recommended structure:

```
tools/corpus/private/
  originals/   # encrypted originals (*.enc) (encrypted-at-rest by the corpus tooling; not Excel “password to open” encryption)
  sanitized/   # sanitized workbooks (plaintext, safe-ish)
  metadata/    # ingest metadata + sanitize options
  reports/     # per-workbook triage reports (json)
```

---

## Privacy controls / anonymization

The sanitization pipeline is implemented in `tools/corpus/sanitize.py` and supports:

- **Redact cell values** while preserving formulas and sheet structure
  - numeric cells are normalized to `0`
  - string cells are replaced with `"REDACTED"`
  - formula cached values are removed (to avoid leaking computed results)
  - formula **string literals** are also redacted/hardened to prevent secrets surviving in formulas
  - pivot cache records/items are cleared (pivot caches can otherwise retain full plaintext copies of source data)
    - See [`docs/21-xlsx-pivots.md`](./21-xlsx-pivots.md) for the relevant OpenXML parts (`xl/pivotCache/*`).
  - conditional formatting + data validation formulas (and their string literals) are hardened as well
  - dialog sheets/macrosheets (and chart sheets) are sanitized like normal worksheets (same cell redaction rules)
- **Hash strings** (`--hash-strings --hash-salt ...`)
  - shared strings / inline strings are replaced with stable `H_<digest>` tokens
  - additional text surfaces are hashed too (comments, headers/footers, drawing text, table names, formula string literals)
  - use a private salt to avoid dictionary attacks
- **Remove external links**
  - drops `xl/externalLinks/**`
  - scrubs any relationship with `TargetMode="External"` to `https://redacted.invalid/`
  - scrubs hyperlink display/tooltips inside worksheets
- **Remove secrets**
  - drops common secret-bearing parts like `xl/connections.xml`, `customXml/**`, and `xl/queryTables/**`
  - removes `xl/vbaProject.bin`/`xl/vbaProjectSignature.bin`, `customUI/**`, and embedded binaries like `xl/media/**`
  - drops other binary/metadata-heavy parts commonly containing hostnames/usernames: `xl/printerSettings/**`, `xl/revisions/**`, `xl/webExtensions/**`, `xl/model/**`
  - removes preview images like `docProps/thumbnail.*`
  - removes `docProps/custom.xml` (custom document properties)
  - strips workbook/sheet protection password hashes and file sharing usernames from `xl/workbook.xml` / worksheets
- **Scrub metadata**
  - redacts author fields in `docProps/core.xml` and sensitive fields in `docProps/app.xml`
  - removes workbook defined names (`<definedNames>`) which often embed business terms
  - scrubs comments (`xl/comments*.xml`), headers/footers, drawing text, and table/table-column names
    - structured references in formulas are rewritten to match sanitized table/column names
  - rewrites/scrubs `docProps/app.xml` `TitlesOfParts` (sheet title metadata)

As a defense-in-depth safety net, `tools/corpus/sanitize.py` also includes a **leak scanner**
(`scan_xlsx_bytes_for_leaks`) that can be used to fail CI if sanitized outputs still match high-risk
patterns (emails, URLs, AWS keys, JWTs) or known plaintext strings.

Sanitized outputs are also written with **deterministic ZIP metadata** (stable timestamps) so re-sanitizing the same
input does not create noisy diffs or leak ingest time.

The design constraint: triage reports must not leak plaintext cell values.

---

## Local usage

### Generate an encryption key

```bash
python tools/corpus/keygen.py
export CORPUS_ENCRYPTION_KEY="..."
```

### Ingest a workbook into the private corpus

```bash
python -m tools.corpus.ingest --input /path/to/workbook.xlsx
```

For additional hardening, consider enabling deterministic sheet renaming:

```bash
python -m tools.corpus.ingest --input /path/to/workbook.xlsx --rename-sheets
```

This stores:
- an **encrypted** original (`originals/*.enc`)
- a sanitized workbook (`sanitized/*.xlsx`)
- ingest metadata (`metadata/*.json`)
- a triage report (`reports/*.json`)

Note: ingest uses triage `--privacy-mode private` behavior when generating the stored triage report
(`reports/*.json`), so filenames/custom domains are not written into that report JSON by default.

### Run triage

```bash
python -m tools.corpus.triage \
  --corpus-dir tools/corpus/public \
  --out-dir tools/corpus/out/public \
  --expectations tools/corpus/public/expectations.json \
  --include-xlsb
```
Note: triage invokes a small Rust helper (built via `cargo`) to run the `formula-xlsx` / `formula-xlsb` round-trip and `xlsx-diff`
structural comparison, so a Rust toolchain must be available.

Triage output layout (`--out-dir`):

- `index.json` – run metadata + an ordered list of reports (`id`, `display_name`, `file`)
- `reports/*.json` – one per workbook; filenames are deterministic and non-colliding:
  `"<sha256[:16]>-<path_hash>.json"` where `path_hash` is a stable hash of the workbook's path
  relative to `--corpus-dir`. This ensures `--jobs 1` and `--jobs N` produce the same report file set.

#### Round-trip diff categorization (`round_trip_failure_kind`)

Triage reports always include a coarse `failure_category` (e.g. `open_error`, `calc_mismatch`,
`render_error`, `round_trip_diff`).

For `failure_category=round_trip_diff`, triage also emits a higher-signal `round_trip_failure_kind`
derived from the Rust helper’s **diff part groups** (privacy-safe: no XML values are emitted; only
OPC part paths + group labels).

Examples:

- `round_trip_rels` – diffs only in `*.rels`
- `round_trip_content_types` – diffs include `[Content_Types].xml`
- `round_trip_styles` – diffs include `xl/styles.xml` (XLSX) or `xl/styles.bin` (XLSB)
- `round_trip_worksheets` – diffs include `xl/worksheets/*`
- `round_trip_shared_strings` – diffs include `xl/sharedStrings.xml` (XLSX) or `xl/sharedStrings.bin` (XLSB)
- `round_trip_media` – diffs include `xl/media/*`
- `round_trip_doc_props` – diffs include `docProps/*`
- `round_trip_workbook` – diffs include `xl/workbook.xml` (XLSX) or `xl/workbook.bin` (XLSB)
- `round_trip_theme` – diffs include `xl/theme/*`
- `round_trip_pivots` – diffs include `xl/pivotTables/*` or `xl/pivotCache/*`
- `round_trip_charts` – diffs include `xl/charts/*`
- `round_trip_drawings` – diffs include `xl/drawings/*`
- `round_trip_tables` – diffs include `xl/tables/*`
- `round_trip_external_links` – diffs include `xl/externalLinks/*`
- `round_trip_other` – anything else

For large private corpora, you can speed up triage by running workbooks in parallel:

```bash
python -m tools.corpus.triage --corpus-dir tools/corpus/private --out-dir tools/corpus/out/private --jobs 4
```

For the **private** corpus (or any environment where triage JSON is uploaded as an artifact), prefer
`--privacy-mode private` to avoid leaking original filenames or custom URI domains.

When using the recommended `originals/` + `sanitized/` layout, triage defaults to preferring sanitized inputs to avoid
double-processing (and to avoid accidentally parsing encrypted originals). Override with
`--input-scope {auto,sanitized,originals,all}` as needed (note: `originals` requires `CORPUS_ENCRYPTION_KEY`).

```bash
python -m tools.corpus.triage \
  --corpus-dir tools/corpus/private \
  --out-dir tools/corpus/out/private \
  --privacy-mode private
```

In `privacy-mode=private`:

- `display_name` is replaced with a stable anonymized value: `workbook-<sha256[:16]>.{xlsx,xlsm,xlsb}`
  (preserves the input extension when known)
- custom URI-like strings (relationship types, namespaces, etc.) are replaced with `sha256=<digest>` unless
  they use standard OpenXML/Microsoft schema hosts (this also redacts expanded XML namespaces embedded in
  diff paths like `{http://example.com/ns}attr`, plus other URL/path-like substrings such as `file:///...`,
  `//server/share/...`, and common absolute filesystem paths)
- non-standard/custom formula function names (e.g. add-in/UDF prefixes) in `functions` are replaced with
  `sha256=<digest>`
- `run_url` is hashed when it points at a non-`github.com` host (e.g. GitHub Enterprise Server domains)
- `index.json` hashes local filesystem paths (`corpus_dir`, `input_dir`) to avoid leaking usernames/mount points
- expectation comparison (when `--expectations` is provided) uses the anonymized `display_name` keys

### Diff policy (ignored parts + calcChain)

The `xlsx-diff` step classifies differences by severity:

- **CRITICAL** – counts toward CI regression gating (`diff_critical_count`)
- **WARN** / **INFO** – surfaced in reports and dashboards, but do not fail CI by default

`tools/corpus/triage.py` ignores a small set of parts that are typically noisy across writers
(`docProps/core.xml`, `docProps/app.xml`).

To run a strict diff that includes those parts (useful for measuring full package stability), pass:

```bash
python -m tools.corpus.triage \
  --corpus-dir tools/corpus/public \
  --out-dir tools/corpus/out/public-strict \
  --no-default-diff-ignore
```

#### Suppressing known-noise XML diffs (ignore-path)

Some Excel writers emit volatile or non-semantic attributes that churn across round-trips (for example
`x14ac:dyDescent` in DrawingML text runs, or `xr:uid` in rich-data extensions). To suppress these without
ignoring the entire part, triage exposes `xlsx-diff`’s fine-grained XML ignore rules:

```bash
# Ignore any XML diff whose path contains the substring (repeatable).
python -m tools.corpus.triage ... \
  --diff-ignore-path dyDescent \
  --diff-ignore-path xr:uid
```

```bash
# Scope an ignore rule to matching parts (repeatable).
python -m tools.corpus.triage ... \
  --diff-ignore-path-in "xl/worksheets/*.xml:xr:uid"
```
#### Ignore presets (known-volatile Excel XML attributes)

Some OOXML attributes are **not stable across Excel saves** (for example `xr:uid` revision identifiers,
`a:blip/@cstate`, or `x14ac:dyDescent`). These can create noisy diffs when comparing a workbook against a
re-saved version.

To opt in to suppressing these known-volatile XML attribute diffs, pass:

```bash
python -m tools.corpus.triage ... --diff-ignore-preset excel-volatile-ids
```

This is intentionally **opt-in** so default corpus triage remains strict.

```bash
# Ignore only diffs of a specific kind (repeatable).
# Format: <kind>:<path_substring>
python -m tools.corpus.triage ... \
  --diff-ignore-path-kind "attribute_changed:@"
```

```bash
# Scope a kind-filtered ignore rule to matching parts (repeatable).
# Format: <part_glob>:<kind>:<path_substring>
python -m tools.corpus.triage ... \
  --diff-ignore-path-kind-in "xl/worksheets/*.xml:attribute_changed:@"
```

To ignore a family of parts by pattern (repeatable), pass `--diff-ignore-glob`:

```bash
# Ignore embedded media assets (images, etc.)
python -m tools.corpus.triage ... --diff-ignore-glob 'xl/media/*'
```

#### calcChain (`xl/calcChain.xml`)

Excel workbooks may include a **calculation chain** (`xl/calcChain.xml`) that records formula dependency order.
Many producers drop or regenerate it during recalculation, so churn is common.

However, preserving calcChain *when possible* is a project goal, so corpus triage **does not ignore calcChain
diffs**. `xlsx-diff` downgrades calcChain-related diffs (including associated relationship / content-type changes)
to **WARN**, so they show up in round-trip metrics and dashboards without breaking CI gates.

To locally hide calcChain noise (restoring the old triage behavior), run triage with:

```bash
python -m tools.corpus.triage ... --diff-ignore xl/calcChain.xml
```

To make calcChain diffs count as **CRITICAL** (strict round-trip preservation scoring), run triage with:

```bash
python -m tools.corpus.triage ... --strict-calc-chain
```

### Isolate round-trip diffs by part ("diff root cause isolator")

When a workbook fails round-trip (`diff_critical_count > 0`), it can be useful to quickly identify **which package
parts** are responsible, without emitting any raw XML/text to logs.

```bash
python -m tools.corpus.minimize --input /path/to/workbook.xlsx
```

If you plan to upload the resulting JSON as an artifact (e.g. for private corpora), prefer:

```bash
python -m tools.corpus.minimize --input /path/to/workbook.xlsx --privacy-mode private
```

This produces a privacy-safe summary:

- `critical_parts`: list of OPC part names with at least one `CRITICAL` diff
- `part_counts`: per-part diff counts (C/W/I/total)
- `rels_critical_ids`: for `.rels` parts, the relationship Ids (`rId...`) involved in critical diffs
- `critical_part_hashes`: sha256 + size of each critical part (helps correlate failures without leaking content)
- `parts_with_diffs`: stable list of all parts with diffs, including a coarse group label (rels/styles/media/etc)

The tool always writes a JSON summary file (default: `tools/corpus/out/minimize/<sha16>.json`).

To include the usually-ignored noisy metadata parts (`docProps/*`) in the diff analysis, pass:

```bash
python -m tools.corpus.minimize --input /path/to/workbook.xlsx --no-default-diff-ignore
```

To also attempt to write a smaller workbook that still reproduces the critical diffs:

```bash
python -m tools.corpus.minimize --input /path/to/workbook.xlsx --out-xlsx minimized.xlsx
```

To fail fast on suspicious plaintext in a corpus directory:

```bash
python -m tools.corpus.triage --corpus-dir tools/corpus/public --out-dir tools/corpus/out/public --leak-scan
```

### Generate the scorecard/dashboard

```bash
python -m tools.corpus.dashboard --triage-dir tools/corpus/out/public
cat tools/corpus/out/public/summary.md
```

For private corpus artifacts, prefer running the dashboard with `--privacy-mode private` as well:

```bash
python -m tools.corpus.dashboard --triage-dir tools/corpus/out/private --privacy-mode private
```

The dashboard is privacy-safe (no raw XML/value diffs) and includes:

- overall pass rates (open/calc/render/round-trip)
- aggregate diff totals (critical/warn/info)
- top diff **parts** contributing to CRITICAL diffs and to diffs overall
- top diff **part groups** (rels/content types/styles/worksheets/etc.) contributing to CRITICAL diffs and to diffs overall
- a **Timings** section (per-step `duration_ms` stats like `p50`/`p90`)
- (when `--recalc` is enabled) aggregate **cell-level calculation fidelity** (`mismatched_cells` / `formula_cells`)

The machine-readable `summary.json` includes these breakdowns as `top_diff_parts_*` and
`top_diff_part_groups_*`, and timings under `timings`, for metrics/dashboarding and CI gates.
Cell-level calc fidelity (when available) is stored under `calculate_cells` as `{formula_cells, mismatched_cells, fidelity}`.

### Append a trend time series entry (machine-readable)

Each corpus dashboard run can also append a compact JSON entry (rates, diff totals, etc.) to a
time-series file:

```bash
python -m tools.corpus.dashboard \
  --triage-dir tools/corpus/out/private \
  --append-trend tools/corpus/out/private/trend.json \
  --trend-max-entries 90
```

Trend entries are intentionally compact (rates + diff totals + a few key size/timing percentiles) so they can
be cached in CI and plotted over time.

Each entry includes a `schema_version` field so downstream tooling can evolve safely as new metrics are added.

Appending is idempotent for a given triage run: if you re-run the dashboard on the same `--triage-dir` and the
last trend entry already has the same `timestamp`, it will be replaced instead of appended (prevents accidental
duplicate points when regenerating dashboards from artifacts).

Key fields currently emitted include:

- `open_rate`, `round_trip_rate`
- `calc_rate` / `render_rate` (rates among attempted workbooks)
- `calc_cell_fidelity` (cell-level calc fidelity across all recalc-checked formula cells)
- `diff_totals.{critical,warning,info,total}`
- `size_overhead_p90` (output/input size ratio p90 for successful round-trips)
- `part_change_ratio_p90` / `part_change_ratio_critical_p90` (privacy-safe churn signal)

To generate a quick Markdown delta between the last two entries:

```bash
python -m tools.corpus.trend_delta --trend-json tools/corpus/out/private/trend.json
```

The scheduled private corpus workflow (`.github/workflows/corpus.yml`) restores/saves this
`trend.json` file via GitHub Actions cache so it grows over time, and uploads it as part of the
private corpus artifact.

For convenience, CI also uploads a small stable-named artifact containing just the latest trend
file:

- `corpus-private-trend` → `tools/corpus/out/private/trend.json`

### Promote a workbook into the public subset

To add a new public, non-sensitive regression fixture (base64 + expectations) in one step:

```bash
python -m tools.corpus.promote_public \
  --input /path/to/sanitized.xlsx \
  --name my-regression-case \
  --confirm-sanitized
```

Or, to sanitize a raw workbook during promotion:

```bash
python -m tools.corpus.promote_public \
  --input /path/to/raw.xlsx \
  --name my-regression-case \
  --sanitize
```

The command:
1) writes `tools/corpus/public/<name>.(xlsx|xlsm|xlsb).b64`,
2) runs triage on the resulting bytes, and
3) updates `tools/corpus/public/expectations.json` (refusing to overwrite unless `--force` is passed).

Optional knobs:

- reduce report size: `--diff-limit 10`
- enable heavier checks (slower): `--recalc --render-smoke`
- refresh expectations for an existing fixture: run with `--force`
- preview without writing files: `--dry-run`

Note: when `--input` already points at the canonical `tools/corpus/public/*.b64` fixture and an expectations entry is
present, the command exits successfully without running triage (no Rust toolchain required). Pass `--force` to re-run
triage and refresh the expectations entry.

`--dry-run` prints a JSON summary including `needs_force` to indicate whether overwriting the fixture/expectations would
require `--force`.

For `.xlsb` inputs, `--sanitize` is not supported; provide an already-sanitized workbook. Leak scanning still runs by
default (XLSB is also an OPC zip container).

Filename safety: if you omit `--name` when promoting from outside `tools/corpus/public/`, the command uses a
hash-based name (`workbook-<sha256[:16]>.{xlsx,xlsm,xlsb}`) instead of the local filename to avoid leaking
customer/org names.

### Generate a unified compatibility scorecard (corpus + Excel-oracle)

The corpus dashboard captures **read + round-trip** behavior, while the Excel-oracle harness captures
**calculation fidelity**. To get a single view across both, generate a unified scorecard:

```bash
# 1) Run corpus triage + dashboard (produces tools/corpus/out/<variant>/summary.json)
python -m tools.corpus.triage --corpus-dir tools/corpus/public --out-dir tools/corpus/out/public \
  --expectations tools/corpus/public/expectations.json
python -m tools.corpus.dashboard --triage-dir tools/corpus/out/public

# 2) Run the Excel-oracle gate (produces tests/compatibility/excel-oracle/reports/mismatch-report.json)
python tools/excel-oracle/compat_gate.py

# 3) Merge into a single markdown scorecard
python tools/compat_scorecard.py --out-md compat_scorecard.md
# or:
python -m tools.compat_scorecard --out-md compat_scorecard.md
```

By default, `tools/compat_scorecard.py` looks for:

- the newest `tools/corpus/out/<variant>/summary.json` (public/private/strict variants)
- `tests/compatibility/excel-oracle/reports/mismatch-report.json`

If one input is missing, it exits non-zero and prints which file is missing (use
`--allow-missing-inputs` to render a partial scorecard).

---

## CI integration

Workflow: `.github/workflows/corpus.yml`

- **PR CI** runs the public subset (`tools/corpus/public/`) and fails on regressions against
  `tools/corpus/public/expectations.json`.
- **Scheduled CI** is intended to run the full private corpus when secrets are available.
- **Compatibility gates (scheduled only)** enforce aggregate targets (e.g. "97%+ round-trip preservation")
  for the private corpus using `tools/corpus/compat_gate.py`.

### Unified compatibility scorecard (CI)

The project tracks:

- L1 Read compatibility (corpus open rate)
- L2 Calculate fidelity (Excel-oracle mismatch rate → pass rate)
- L4 Round-trip preservation (corpus round-trip rate)

To produce a single view in CI, the repo includes an aggregator workflow:

- `.github/workflows/compat-scorecard.yml` (trigger: `workflow_run` of `corpus` and `Excel Compatibility (Oracle)`)

This workflow downloads the two summary artifacts (by matching `head_sha`), runs
`python tools/compat_scorecard.py --allow-missing-inputs`, uploads a `compat-scorecard` artifact, and appends the
markdown to the job summary.

You can also run the workflow manually via `workflow_dispatch` and optionally provide a target commit SHA to
backfill/debug scorecards for older runs.

Corpus summary artifacts:

- `corpus-public-summary` → `tools/corpus/out/public/summary.json` + `tools/corpus/out/public/summary.md`
- `corpus-private-summary` → `tools/corpus/out/private/summary.json` + `tools/corpus/out/private/summary.md`

Excel-oracle artifacts:

- `excel-oracle-summary` → `tests/compatibility/excel-oracle/reports/mismatch-report.json` + `summary.md`
- `excel-oracle-artifacts` → full datasets + reports (useful for deep debugging)

### Workflow dispatch knobs (Calculate/Render coverage)

The `corpus` workflow supports `workflow_dispatch` inputs so you can opt into heavier checks without code
changes:

- `recalc`: run `tools.corpus.triage --recalc` (Calculate / L2). The nightly scheduled private corpus run
  enables this by default.
- `render_smoke`: run `tools.corpus.triage --render-smoke` (Render smoke / L3).
- `min_calc_rate`: optional CI gate threshold for Calculate pass rate (among attempted workbooks).
- `min_calc_cell_fidelity`: optional CI gate threshold for Calculate **cell-level fidelity** (among attempted formula cells).
- `min_render_rate`: optional CI gate threshold for Render pass rate (among attempted workbooks).

Note: `--recalc` compares engine results against **cached formula values stored in the workbook**. If your
corpus inputs were aggressively sanitized (e.g. redacting cell values), cached formula results are often
removed, which will cause the calculate step to show `SKIP` with `0 attempted` workbooks. In that case,
run triage against inputs that still contain cached values (commonly the encrypted `originals/` corpus).

In CI, the scheduled private corpus workflow will automatically prefer `--input-scope originals` when
`recalc` is enabled and an `originals/` directory is present.

These inputs apply to both the public and private corpus jobs; the private job will still skip if the
required secrets/corpus archive are not configured.

For scheduled runs, you can also opt into nightly render smoke by setting the repo variable
`CORPUS_RUN_RENDER_SMOKE=true` (see `.github/workflows/corpus.yml`).

By default the scheduled private corpus run enables `recalc` (Calculate / L2). To disable it, set the
repo variable `CORPUS_RUN_RECALC=false`.

### Compatibility gate thresholds

After generating `tools/corpus/out/<variant>/summary.json` via `tools/corpus/dashboard.py`, you can enforce minimum
pass rates:

```bash
python -m tools.corpus.compat_gate \
  --triage-dir tools/corpus/out/private \
  --min-round-trip-rate 0.97
```

Supported thresholds:

- `--min-open-rate`
- `--min-round-trip-rate`
- `--min-calc-rate` *(when triage is run with `--recalc`; measured among workbooks where calculation was attempted)*
- `--min-calc-cell-fidelity` *(when triage is run with `--recalc`; cell-level fidelity across attempted formula cells, requires cached values)*
- `--min-render-rate` *(when triage is run with `--render-smoke`; measured among workbooks where render was attempted)*

### Optional performance gates (scheduled/private)

The dashboard also supports opt-in **performance regression gates** on p90 step durations:

- `--gate-load-p90-ms <ms>`
- `--gate-round-trip-p90-ms <ms>`

In CI, these are intended to be enabled for the private corpus workflow (via workflow dispatch inputs
or repo variables; see `.github/workflows/corpus.yml`). They are **off by default** so PR CI behavior
does not change.

### Expected secrets (scheduled CI)

Recommended secrets for a real deployment:

- `CORPUS_ENCRYPTION_KEY` – Fernet key used to decrypt `*.enc` originals.
- `CORPUS_PRIVATE_TAR_B64` – base64-encoded tarball containing `tools/corpus/private/`.

The scheduled job should:
- download/decode the corpus archive
- run triage + dashboard generation
- upload results as artifacts (and/or publish to a dashboard)

---

## Extending triage (future work)

`tools/corpus/triage.py` is intended to be a compatibility regression harness that runs:

- load (via `formula-xlsx`)
- load (via `formula-xlsx` / `formula-xlsb`)
- **optional** recalculation correctness checks (`--recalc`)
- **optional** headless render/print smoke (`--render-smoke`)
- round-trip save (via `formula-xlsx`)
- round-trip save (via `formula-xlsx` / `formula-xlsb`)
- structural diff (via `xlsx-diff`)

Recalc/render are opt-in because they are heavier and may exercise engine coverage gaps; the scheduled private
corpus job is the recommended place to enable them.
