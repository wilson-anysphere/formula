# Public corpus subset

This directory contains a **small, non-sensitive** corpus subset that can run in PR CI.

Files may be stored as:
- `*.xlsx` / `*.xlsm` (binary) *(not recommended for git)*
- `*.xlsx.b64` / `*.xlsm.b64` (base64 text) *(preferred for tiny fixtures)*

The workflow `.github/workflows/corpus.yml` runs triage + dashboard generation against this directory.

## Adding a new public fixture

Use the helper to avoid manual base64 encoding + editing `expectations.json`:

```bash
# Promote an *already sanitized* workbook (recommended workflow when sourced from the private corpus):
python -m tools.corpus.promote_public \
  --input /path/to/sanitized.xlsx \
  --name my-regression-case \
  --confirm-sanitized

# Or: sanitize during promotion (for raw inputs), then leak-scan by default:
python -m tools.corpus.promote_public \
  --input /path/to/raw.xlsx \
  --name my-regression-case \
  --sanitize
```

This will:

- write `tools/corpus/public/<name>.xlsx.b64`
- run triage on it (report written under `tools/corpus/out/promote-public/`)
- upsert `tools/corpus/public/expectations.json`

By default the command refuses to overwrite existing fixtures/expectations; re-run with `--force` if needed.
