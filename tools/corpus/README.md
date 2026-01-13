# Compatibility corpus tooling (`tools/corpus`)

This directory contains a privacy-safe pipeline for building and triaging a "real-world XLSX zoo".

Most users should start with:

- `python -m tools.corpus.ingest` – ingest a workbook into a private corpus (stores encrypted original + sanitized copy)
- `python -m tools.corpus.triage` – run automated triage over a corpus directory
- `python -m tools.corpus.promote_public` – promote a (sanitized) workbook into the public subset (`tools/corpus/public/`)
- `python -m tools.corpus.dashboard` – generate a scorecard from triage reports (optionally `--append-trend .../trend.json`)
- `python -m tools.corpus.minimize` – summarize which workbook parts are responsible for round-trip diffs (privacy-safe); can optionally emit a minimized workbook via `--out-xlsx`

For full docs, see: `docs/compatibility-corpus.md`.
