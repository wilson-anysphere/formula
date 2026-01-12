# Running Excel COM automation in CI (self-hosted runner)

The Excel oracle harness (`tools/excel-oracle/run-excel-oracle.ps1`) requires **real Microsoft Excel desktop**.

GitHub-hosted `windows-latest` runners typically do **not** include Excel, so generating oracle datasets in CI usually requires a **self-hosted Windows runner** with Office installed.

## High-level setup

1) Provision a Windows machine or VM (Windows 11 / Windows Server).

2) Install Microsoft Office / Microsoft Excel.
   - Ensure licensing is compliant for CI usage.

3) Install the GitHub Actions self-hosted runner for this repo.
   - Recommended: assign labels like `windows`, `excel`, `office`.

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

## Common pitfalls

- **Runner as a service:** Excel COM may fail or hang when no desktop session is available.
- **First-run prompts:** Excel may show first-run UI / privacy prompts. Launch Excel once manually under the runner user to clear them.
- **Popups/alerts:** The harness sets `DisplayAlerts = false`, but other dialogs can still appear if Excel isn't fully configured.
- **Locale:** The harness uses `Range.Formula2` / `Range.Formula` (English) rather than `FormulaLocal`, which is usually safer across locales.
- **Number separators:** The harness forces US-style decimal/thousands separators (`.` / `,`) so text→number coercion cases like `"1,234"` are deterministic across runner locales.
