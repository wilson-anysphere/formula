# Running Excel COM automation in CI (self-hosted runner)

The Excel oracle harness (`tools/excel-oracle/run-excel-oracle.ps1`) requires **real Microsoft Excel desktop**.

GitHub-hosted `windows-latest` runners typically do **not** include Excel, so generating oracle datasets in CI usually requires a **self-hosted Windows runner** with Office installed.

## High-level setup

1) Provision a Windows machine or VM (Windows 11 / Windows Server).

2) Install Microsoft Office / Microsoft Excel.
   - Ensure licensing is compliant for CI usage.

3) Install the GitHub Actions self-hosted runner for this repo.
   - Recommended: assign labels like `windows`, `excel`, `office`.
   - The workflow `.github/workflows/excel-compat.yml` routes Excel-generation runs to a runner
     labeled `self-hosted` + `windows` + `excel`.

4) Run the runner **interactively** (important).
   - Excel COM automation often fails when the runner is installed as a Windows Service because Office expects a user profile + desktop session.
   - Prefer launching `run.cmd` from an interactive login session for the runner account.

5) Trigger the workflow manually:
   - Workflow: **Excel Compatibility (Oracle)**
   - Inputs: `mode=generate-oracle` (or `validate-engine`)
   - Inputs: `oracle_source=generate`

6) Download artifacts and pin the dataset:

```bash
python tools/excel-oracle/pin_dataset.py \
  --dataset tests/compatibility/excel-oracle/datasets/excel-oracle.json \
  --pinned tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json \
  --versioned-dir tests/compatibility/excel-oracle/datasets/versioned
```

Commit the pinned dataset to enable PR/push validation without Excel.

## Incremental patching (recommended for targeted parity work)

If you only need to validate/pin **a subset** of the corpus (e.g. odd-coupon bonds) you do *not*
need to regenerate the entire dataset.

Instead, run the convenience wrapper that:

1) runs Excel on a subset corpus, then
2) overwrites just those `caseId`s in the pinned dataset (merge-friendly).

Example (odd-coupon boundary equality cases):

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/patch-pinned-dataset-with-excel.ps1 `
  -SubsetCasesPath tools/excel-oracle/odd_coupon_boundary_cases.json
```

This workflow records compact provenance under `source.patches` in the pinned dataset so it’s clear
which Excel build produced the patched values.

Then verify parity:

```bash
python tools/excel-oracle/compat_gate.py --include-tag odd_coupon
```

## Common pitfalls

- **Runner as a service:** Excel COM may fail or hang when no desktop session is available.
- **First-run prompts:** Excel may show first-run UI / privacy prompts. Launch Excel once manually under the runner user to clear them.
- **Popups/alerts:** The harness sets `DisplayAlerts = false`, but other dialogs can still appear if Excel isn't fully configured.
- **Locale:** The harness uses `Range.Formula2` / `Range.Formula` (English) rather than `FormulaLocal`, which is usually safer across locales.
- **Number separators:** The harness forces US-style decimal/thousands separators (`.` / `,`) so text→number coercion cases like `"1,234"` are deterministic across runner locales.

## Generating locale translation sources on a self-hosted runner (optional)

The same self-hosted Windows runner can also be used to extract **locale function-name translations**
from real Excel (useful for keeping `de-DE` / `es-ES` sources complete when new functions are added
to `shared/functionCatalog.json`).

From repo root on the runner:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
  -LocaleId de-DE `
  -OutPath crates/formula-engine/src/locale/data/sources/de-DE.json

node scripts/generate-locale-function-tsv.js
node scripts/generate-locale-function-tsv.js --check
```

Important: the extracted spellings reflect the **active Excel UI language**. Install the relevant
Office language pack and set Excel's display language before extracting.
