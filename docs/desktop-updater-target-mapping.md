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
- `linux-aarch64`

If Tauri changes these identifiers in a future upgrade, our CI guardrail
(`scripts/verify-tauri-latest-json.mjs`, which wraps `scripts/ci/validate-updater-manifest.mjs`) is
expected to fail loudly with a clear “expected vs actual” diff, and this document should be updated
alongside the validator.

### Common equivalents (not accepted by tagged-release CI)

The CI validator and the standalone checker are intentionally **strict** about the `latest.json`
platform key names so toolchain changes fail loudly.

When inspecting manifests from **local builds** or ad-hoc tooling, you may see alternate platform
key spellings (often Rust target triples). Treat these as equivalent to the canonical keys above
for debugging, but note that tagged-release CI is intentionally strict and expects the canonical
keys.

- `universal-apple-darwin` → `darwin-universal`
- `x86_64-pc-windows-msvc` → `windows-x86_64`
- `windows-arm64` / `aarch64-pc-windows-msvc` → `windows-aarch64`
- `x86_64-unknown-linux-gnu` → `linux-x86_64`
- `aarch64-unknown-linux-gnu` → `linux-aarch64`

If a tagged release ever ships with a different key spelling, update both this document and the CI
validator together.

## Mapping table (build target → platform key → updater artifact)

The updater does **not** necessarily use the same artifact you’d download for a manual install.
The table below documents what each platform key should point to in `latest.json`.

| OS / Arch | Build target (Tauri `--target`) | `latest.json` platform key | Updater asset type (`platforms[key].url`) |
| --- | --- | --- | --- |
| macOS universal (Intel + Apple Silicon) | `universal-apple-darwin` | `darwin-universal` | `*.app.tar.gz` (preferred) or `*.tar.gz` (updater archive; **not** the `.dmg`) |
| Windows x64 | `x86_64-pc-windows-msvc` | `windows-x86_64` | `*.msi` (preferred) or `*.exe` |
| Windows ARM64 | `aarch64-pc-windows-msvc` | `windows-aarch64` | `*.msi` (preferred) or `*.exe` |
| Linux x86_64 | `x86_64-unknown-linux-gnu` | `linux-x86_64` | `*.AppImage` (self-updatable; **not** `.deb`/`.rpm`) |
| Linux ARM64 | `aarch64-unknown-linux-gnu` | `linux-aarch64` | `*.AppImage` (self-updatable; **not** `.deb`/`.rpm`) |

### Notes

- `latest.json` may contain other metadata (`notes`, `pub_date`, etc). Those fields are not treated
  as stable for CI verification.
- Formula still publishes **additional** artifacts (DMG, NSIS `.exe`, `.deb`, `.rpm`) for user
  convenience; the updater keys above intentionally validate the *updatable* artifact.
- Windows distribution requirement: the updater uses whichever installer is referenced in
  `latest.json` (`.msi` or `.exe`; in this repo it is typically `.msi`). Tagged releases must also
  ship a corresponding **NSIS `.exe` installer** for both `windows-x86_64` and `windows-aarch64`. The
  release workflow enforces that both `.msi` and `.exe` assets exist per architecture.
- Windows multi-arch safety requirement: the `.msi` / `.exe` filenames must include an explicit
  architecture marker (for example `x64`, `x86_64`, `amd64`, `win64` vs `arm64`, `aarch64`) so that
  a multi-target run cannot overwrite/clobber assets on the GitHub Release. CI enforces this via:
  - `scripts/ci/validate-updater-manifest.mjs` (ensures `latest.json` points at arch-specific assets)
  - `verify-release-assets` in `.github/workflows/release.yml` (ensures uploaded assets are uniquely named)
- To inspect a manifest locally:
  - `jq '.platforms | keys' latest.json`
  - `jq -r '.platforms["windows-aarch64"].url' latest.json`

- To validate a downloaded manifest signature offline:
  - `node scripts/ci/verify-updater-manifest-signature.mjs latest.json latest.json.sig`
