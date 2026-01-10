# Public corpus subset

This directory contains a **small, non-sensitive** corpus subset that can run in PR CI.

Files may be stored as:
- `*.xlsx` / `*.xlsm` (binary) *(not recommended for git)*
- `*.xlsx.b64` / `*.xlsm.b64` (base64 text) *(preferred for tiny fixtures)*

The workflow `.github/workflows/corpus.yml` runs triage + dashboard generation against this directory.

