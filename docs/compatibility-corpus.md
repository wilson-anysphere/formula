# Compatibility corpus ("real-world XLSX zoo")

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

### Run triage

```bash
python -m tools.corpus.triage \
  --corpus-dir tools/corpus/public \
  --out-dir tools/corpus/out/public \
  --expectations tools/corpus/public/expectations.json \
  --include-xlsb
```
Note: triage invokes a small Rust helper (built via `cargo`) to run the `formula-xlsx` round-trip and `xlsx-diff`
structural comparison, so a Rust toolchain must be available.

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
  diff paths like `{http://example.com/ns}attr`)
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

### Isolate round-trip diffs by part ("diff root cause isolator")

When a workbook fails round-trip (`diff_critical_count > 0`), it can be useful to quickly identify **which package
parts** are responsible, without emitting any raw XML/text to logs.

```bash
python -m tools.corpus.minimize --input /path/to/workbook.xlsx
```

This produces a privacy-safe summary:

- `critical_parts`: list of OPC part names with at least one `CRITICAL` diff
- `part_counts`: per-part diff counts (C/W/I/total)
- `rels_critical_ids`: for `.rels` parts, the relationship Ids (`rId...`) involved in critical diffs
- `critical_part_hashes`: sha256 + size of each critical part (helps correlate failures without leaking content)

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

The dashboard is privacy-safe (no raw XML/value diffs) and includes:

- overall pass rates (open/calc/render/round-trip)
- aggregate diff totals (critical/warn/info)
- top diff **parts** contributing to CRITICAL diffs and to diffs overall
- top diff **part groups** (rels/content types/styles/worksheets/etc.) contributing to CRITICAL diffs and to diffs overall
- a **Timings** section (per-step `duration_ms` stats like `p50`/`p90`)

The machine-readable `summary.json` includes these breakdowns as `top_diff_parts_*` and
`top_diff_part_groups_*`, and timings under `timings`, for metrics/dashboarding and CI gates.

### Append a trend time series entry (machine-readable)

Each corpus dashboard run can also append a compact JSON entry (rates, diff totals, etc.) to a
time-series file:

```bash
python -m tools.corpus.dashboard \
  --triage-dir tools/corpus/out/private \
  --append-trend tools/corpus/out/private/trend.json
```

Trend entries are intentionally compact (rates + diff totals + a few key size/timing percentiles) so they can
be cached in CI and plotted over time.

The scheduled private corpus workflow (`.github/workflows/corpus.yml`) restores/saves this
`trend.json` file via GitHub Actions cache so it grows over time, and uploads it as part of the
private corpus artifact.

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
1) writes `tools/corpus/public/<name>.xlsx.b64`,
2) runs triage on the resulting bytes, and
3) updates `tools/corpus/public/expectations.json` (refusing to overwrite unless `--force` is passed).

Optional knobs:

- reduce report size: `--diff-limit 10`
- enable heavier checks (slower): `--recalc --render-smoke`

### Generate a unified compatibility scorecard (corpus + Excel-oracle)

The corpus dashboard captures **read + round-trip** behavior, while the Excel-oracle harness captures
**calculation fidelity**. To get a single view across both, generate a unified scorecard:

```bash
# 1) Run corpus triage + dashboard (produces tools/corpus/out/**/summary.json)
python -m tools.corpus.triage --corpus-dir tools/corpus/public --out-dir tools/corpus/out/public \
  --expectations tools/corpus/public/expectations.json
python -m tools.corpus.dashboard --triage-dir tools/corpus/out/public

# 2) Run the Excel-oracle gate (produces tests/compatibility/excel-oracle/reports/mismatch-report.json)
python tools/excel-oracle/compat_gate.py

# 3) Merge into a single markdown scorecard
python tools/compat_scorecard.py --out-md compat_scorecard.md
```

By default, `tools/compat_scorecard.py` looks for:

- `tools/corpus/out/**/summary.json` (prefers `tools/corpus/out/public/summary.json` when present)
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

### Workflow dispatch knobs (Calculate/Render coverage)

The `corpus` workflow supports `workflow_dispatch` inputs so you can opt into heavier checks without code
changes:

- `recalc`: run `tools.corpus.triage --recalc` (Calculate / L2). The nightly scheduled private corpus run
  enables this by default.
- `render_smoke`: run `tools.corpus.triage --render-smoke` (Render smoke / L3).

These inputs apply to both the public and private corpus jobs; the private job will still skip if the
required secrets/corpus archive are not configured.

### Compatibility gate thresholds

After generating `tools/corpus/out/**/summary.json` via `tools/corpus/dashboard.py`, you can enforce minimum
pass rates:

```bash
python -m tools.corpus.compat_gate \
  --triage-dir tools/corpus/out/private \
  --min-round-trip-rate 0.97
```

Supported thresholds:

- `--min-open-rate`
- `--min-round-trip-rate`
- `--min-calc-rate` *(when triage is run with `--recalc`)*
- `--min-render-rate` *(when triage is run with `--render-smoke`)*

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
