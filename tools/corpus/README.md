# Compatibility corpus tooling (`tools/corpus`)

This directory contains a privacy-safe pipeline for building and triaging a "real-world spreadsheet zoo"
(`.xlsx`/`.xlsm`/`.xlsb`).

Most users should start with:

- `python -m tools.corpus.ingest` – ingest a workbook into a private corpus (stores encrypted original + sanitized copy)
- `python -m tools.corpus.triage` – run automated triage over a corpus directory
- `python -m tools.corpus.promote_public` – promote a (sanitized) workbook into the public subset (`tools/corpus/public/`)
- `python -m tools.corpus.dashboard` – generate a scorecard from triage reports (optionally `--append-trend .../trend.json`)
- `python tools/compat_scorecard.py` *(or `python -m tools.compat_scorecard`)* – merge the corpus scorecard with the Excel-oracle mismatch report into a single compatibility table
- `python -m tools.corpus.minimize` – summarize which workbook parts are responsible for round-trip diffs (privacy-safe, includes `round_trip_failure_kind`); can optionally emit a minimized workbook via `--out-xlsx`

`promote_public` notes:

- If `--name` is omitted when promoting from outside `tools/corpus/public/`, a hash-based name is used by default
  (`workbook-<sha256[:16]>.{xlsx,xlsm,xlsb}`) to avoid leaking local/customer filenames.
- Re-run with `--force` to refresh an existing fixture's expectations against the current engine behavior.

The dashboard emits both human- and machine-readable outputs:

- `summary.md` (includes a **Timings** section with per-step `duration_ms` stats like p50/p90)
- `summary.json` (includes the same data under `timings`)
- when `--recalc` is enabled, `summary.json` also includes aggregate cell-level calc fidelity under
  `calculate_cells` (mismatches/formula cells)

For round-trip failures (`failure_category=round_trip_diff`), triage reports also include a more
actionable `round_trip_failure_kind` (based on diff part groups like rels/content-types/styles/worksheets,
plus common “other” buckets like workbook/theme/pivots), and the dashboard summarizes these under
`failures_by_round_trip_failure_kind`.

The dashboard also supports opt-in perf regression gates:

- `--gate-load-p90-ms <ms>`
- `--gate-round-trip-p90-ms <ms>`

## Private corpus note

For private corpora following the recommended `originals/` + `sanitized/` layout, `tools.corpus.triage` defaults to
preferring `sanitized/` when present (to avoid double-processing originals). Override with
`--input-scope {auto,sanitized,originals,all}` as needed.

For large corpora, pass `--jobs N` (default `1`) to run per-workbook triage in parallel.

Office-encrypted workbook note (Excel “Encrypt with Password” / `EncryptionInfo` + `EncryptedPackage`):

- By default, Office-encrypted `.xlsx`/`.xlsm`/`.xlsb` inputs are **skipped**.
- To triage Office-encrypted workbooks, pass a password (recommended: via `--password-file` so it
  doesn’t appear in process args):
- `tools.corpus.minimize` supports the same `--password` / `--password-file` options when inspecting
  a single workbook.

```bash
python -m tools.corpus.triage \
  --corpus-dir tools/corpus/private \
  --out-dir tools/corpus/out/private \
  --privacy-mode private \
  --password-file /path/to/workbook-password.txt
```

When triage outputs are uploaded as artifacts (e.g. scheduled CI runs), use `--privacy-mode private` to avoid leaking:

- original filenames (`display_name` is anonymized to `workbook-<sha256[:16]>.{xlsx,xlsm,xlsb}`)
- custom URI domains/paths in relationship/content types and diff paths (hashed as `sha256=<digest>`)
- non-standard/custom formula function names (e.g. add-in/UDF prefixes) in `functions` (hashed as `sha256=<digest>`)
- GitHub Enterprise `run_url` hostnames and local filesystem paths in `index.json`

For private dashboards, run `tools.corpus.dashboard` with `--privacy-mode private` as well. The diff minimizer
supports the same flag (`tools.corpus.minimize --privacy-mode private`) for sharing privacy-safe summaries.

For full docs, see: `docs/compatibility-corpus.md`.

CI note: the scheduled private corpus workflow uploads a stable `corpus-private-trend` artifact
containing `tools/corpus/out/private/trend.json` for easy consumption.

Trend retention note: `tools.corpus.dashboard --append-trend` caps the file to the last 90 entries by
default. Override with `--trend-max-entries N` (set `0` for unlimited).

Append idempotence note: if the last entry in `trend.json` already has the same `timestamp` as the current
run, it is replaced instead of appended. This prevents duplicate points when regenerating dashboards from an
existing triage output directory.

To print a Markdown delta between the last two trend entries:

```bash
python -m tools.corpus.trend_delta --trend-json tools/corpus/out/private/trend.json
```
