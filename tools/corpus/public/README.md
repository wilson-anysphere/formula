# Public corpus subset

This directory contains a **small, non-sensitive** corpus subset that can run in PR CI.

Files may be stored as:
- `*.xlsx` / `*.xlsm` / `*.xlsb` (binary) *(not recommended for git)*
- `*.xlsx.b64` / `*.xlsm.b64` / `*.xlsb.b64` (base64 text) *(preferred for tiny fixtures)*

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

Notes:

- `--sanitize` currently supports `.xlsx` / `.xlsm` inputs only.
- `.xlsb` inputs require `--confirm-sanitized` (leak scanning is ZIP/XLSX based).

This will:

- write `tools/corpus/public/<name>.(xlsx|xlsm|xlsb).b64`
- run triage on it (report written under `tools/corpus/out/promote-public/`)
- upsert `tools/corpus/public/expectations.json`

By default the command refuses to overwrite existing fixtures/expectations; re-run with `--force` if needed.

Note: the command runs corpus triage (same as `python -m tools.corpus.triage`), which builds/executes a small Rust
helper binary. A Rust toolchain is required locally.

If you run the command on an existing `tools/corpus/public/*.b64` fixture that already has an expectations entry,
it will **skip triage** and exit successfully (no Rust required). Use `--force` to refresh expectations.

Tip: you can reduce triage report size by lowering the diff entry cap:

```bash
python -m tools.corpus.promote_public --input ... --diff-limit 10
```

Optional heavier triage checks:

```bash
python -m tools.corpus.promote_public --input ... --recalc --render-smoke
```

Refreshing an existing fixture:

```bash
# Re-run triage and update the expectations entry if the current engine behavior changed.
python -m tools.corpus.promote_public --input tools/corpus/public/simple.xlsx.b64 --force
```
