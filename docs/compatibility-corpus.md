# Compatibility corpus ("real-world XLSX zoo")

## Why this exists

Excel compatibility is won and lost on **real workbooks**. The goal of the compatibility corpus is to:

- collect problematic/representative `.xlsx`/`.xlsm` files (from users and internal sources)
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
- Expected pass/fail states are tracked in `tools/corpus/public/expectations.json`.

### 2) Private corpus (runs in scheduled CI)

`tools/corpus/private/` *(gitignored)*

Populated locally via `tools/corpus/ingest.py`, or in CI by downloading a private archive.

Recommended structure:

```
tools/corpus/private/
  originals/   # encrypted originals (*.enc)
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
  - removes preview images like `docProps/thumbnail.*`
- **Scrub metadata**
  - redacts author fields in `docProps/core.xml` and sensitive fields in `docProps/app.xml`
  - removes `docProps/custom.xml` (custom document properties)
  - removes workbook defined names (`<definedNames>`) which often embed business terms
  - scrubs comments (`xl/comments*.xml`), headers/footers, drawing text, and table/table-column names

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
  --expectations tools/corpus/public/expectations.json
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

---

## CI integration

Workflow: `.github/workflows/corpus.yml`

- **PR CI** runs the public subset (`tools/corpus/public/`) and fails on regressions against
  `tools/corpus/public/expectations.json`.
- **Scheduled CI** is intended to run the full private corpus when secrets are available.

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

The current triage runner includes placeholders for:
- recalculation correctness (`calc mismatch`)
- headless rendering smoke tests
- real round-trip save using the Formula XLSX writer and structural diff (Task 12/90)

As the XLSX read/write/diff tooling lands, wire it into `tools/corpus/triage.py` so the corpus becomes a true
compatibility regression harness.
