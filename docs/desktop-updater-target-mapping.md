# Desktop updater target ↔ `latest.json` platform key mapping

Tauri’s updater manifest (`latest.json`) contains a `platforms` object whose **keys** must match the
runtime `{{target}}` string computed by the updater plugin. When we ship **multi-arch** builds
(macOS universal + Windows ARM64), the exact key names become important for:

- verifying `latest.json` is complete (no “last writer wins” overwrite),
- verifying each platform key points at the **self-updatable** artifact type,
- and doing reliable release-asset verification in CI.

This document is the **source of truth** for the key names we expect in `latest.json` for Formula’s
release workflow.

## Expected `latest.json.platforms` keys (multi-arch release)

As of **Tauri v2.9** + **tauri-action v0.6.1** (see `.github/workflows/release.yml`), a full desktop
release contains **exactly** these `latest.json.platforms` keys:

- `darwin-universal`
- `windows-x86_64`
- `windows-aarch64`
- `linux-x86_64`

If Tauri changes these identifiers in a future upgrade, our CI guardrail
(`scripts/verify-tauri-latest-json.mjs`, which wraps `scripts/ci/validate-updater-manifest.mjs`) is
expected to fail (or require adding a new alias) with a clear “expected vs actual” diff, and this
document should be updated alongside the validator.

### Accepted aliases (CI)

The CI validator accepts a few equivalent key spellings to be resilient to toolchain differences:

- `universal-apple-darwin` → `darwin-universal`
- `x86_64-pc-windows-msvc` → `windows-x86_64`
- `windows-arm64` / `aarch64-pc-windows-msvc` → `windows-aarch64`
- `x86_64-unknown-linux-gnu` → `linux-x86_64`

These are treated as aliases for validation, but the canonical keys above are what we expect
`tauri-action` to produce for Formula releases.

## Mapping table (build target → platform key → updater artifact)

The updater does **not** necessarily use the same artifact you’d download for a manual install.
The table below documents what each platform key should point to in `latest.json`.

| OS / Arch | Build target (Tauri `--target`) | `latest.json` platform key | Updater asset type (`platforms[key].url`) |
| --- | --- | --- | --- |
| macOS universal (Intel + Apple Silicon) | `universal-apple-darwin` | `darwin-universal` | `*.app.tar.gz` (preferred) or `*.tar.gz` (updater archive; **not** the `.dmg`) |
| Windows x64 | `x86_64-pc-windows-msvc` | `windows-x86_64` | `*.msi` (Windows Installer; updater runs this) |
| Windows ARM64 | `aarch64-pc-windows-msvc` | `windows-aarch64` | `*.msi` (Windows Installer; updater runs this) |
| Linux x86_64 | `x86_64-unknown-linux-gnu` | `linux-x86_64` | `*.AppImage` (self-updatable; **not** `.deb`/`.rpm`) |

### Notes

- `latest.json` may contain other metadata (`notes`, `pub_date`, etc). Those fields are not treated
  as stable for CI verification.
- Formula still publishes **additional** artifacts (DMG, NSIS `.exe`, `.deb`, `.rpm`) for user
  convenience; the updater keys above intentionally validate the *updatable* artifact.
- Windows distribution requirement: even though the updater uses the `.msi` for self-update, tagged
  releases must also ship a corresponding **NSIS `.exe` installer** for both `windows-x86_64` and
  `windows-aarch64`. The release workflow enforces that both `.msi` and `.exe` assets exist per
  architecture.
- To inspect a manifest locally:
  - `jq '.platforms | keys' latest.json`
  - `jq -r '.platforms["windows-aarch64"].url' latest.json`

- To validate a downloaded manifest signature offline:
  - `node scripts/ci/verify-updater-manifest-signature.mjs latest.json latest.json.sig`
