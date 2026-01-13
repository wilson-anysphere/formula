# Desktop release process

This repository ships the desktop app via GitHub Releases and Tauri's built-in updater.
Tagged pushes (`vX.Y.Z`) trigger a GitHub Actions workflow that builds installers/bundles for
macOS (Apple Silicon + Intel), Windows, and Linux and uploads them to a **draft** GitHub Release.

Platform/architecture expectations for a release:

- **macOS:** **universal** build (Intel + Apple Silicon): `.dmg` (installer) + `.app.tar.gz` (updater payload).
- **Windows:** **x64** + **ARM64**: installers in both formats (`.msi` + `.exe`) for each architecture.
- **Linux:** `.AppImage` + `.deb` + `.rpm`.

The workflow also uploads updater metadata (`latest.json` + `latest.json.sig`) used by the Tauri
updater.

CI also runs a lightweight **Linux AppImage smoke test** (no GUI) to catch common packaging issues
early:

- Extract the generated `*.AppImage` via `--appimage-extract` (no FUSE required)
- Verify the extracted main ELF binary architecture via `readelf -h`
- Run `ldd` and fail the workflow if any shared libraries are `not found`

See `scripts/ci/check-appimage.sh`.

For the **exact** `latest.json.platforms` key names (and which asset each key should point to),
see:

- `docs/desktop-updater-target-mapping.md`

## Rust toolchain pinning (release stability)

Desktop packaging is sensitive to Rust/toolchain changes. This repo pins Rust via
`rust-toolchain.toml` at the repo root, and CI/release workflows enforce that they install the
same version.

To upgrade Rust, open a PR that bumps `rust-toolchain.toml` (and any workflow pins/comments that the
CI guard requests) and rely on CI to validate the new toolchain before tagging a release.

## Testing the release pipeline (workflow_dispatch)

To test packaging/signing changes without creating a git tag, run the **Desktop Release** workflow
manually from GitHub Actions:

1. Go to **Actions → Desktop Release → Run workflow**.
2. Select the branch/commit you want to test.
3. Leave **upload** unchecked (default). This is **dry-run** mode:
   - bundles are built for all OS/targets
   - outputs are uploaded as **workflow artifacts** (no GitHub Release is created/modified)
4. (Recommended) Set **version** to label the artifacts (example: `0.2.3-test`). You can also set
   **tag** (example: `v0.2.3-test`); if provided, it takes precedence.

If you set **upload=true**, the workflow will create/update a **draft** GitHub Release and attach
assets (matching the tag-driven behavior). This requires providing either `tag` or `version`.

Note: `upload=true` runs the same version validation as a tagged release (the tag/version must
match `apps/desktop/src-tauri/tauri.conf.json`), so for ad-hoc pipeline tests without bumping the
app version, prefer `upload=false`.

## Preflight validations (CI enforced)

The release workflow runs a couple of lightweight preflight scripts before it spends time building
bundles. These checks will fail the release workflow on a tagged push if the repo is not in a
releasable state.

Run them locally from the repo root:

```bash
# Note: CI/release workflows run these scripts under Node 22 (see release.yml).
# Using the same major locally reduces "works locally, breaks in release" drift.
# Ensures the pinned TAURI_CLI_VERSION in .github/workflows/release.yml matches the Tauri crate major/minor.
node scripts/ci/check-tauri-cli-version.mjs

# Ensures the tag version matches apps/desktop/src-tauri/tauri.conf.json "version".
node scripts/check-desktop-version.mjs vX.Y.Z

# Ensures plugins.updater.pubkey/endpoints are not placeholders when the updater is active.
node scripts/check-updater-config.mjs

# Ensures Windows installers will install WebView2 if it is missing.
node scripts/ci/check-webview2-install-mode.mjs

# Ensures Windows Authenticode timestamping uses HTTPS.
node scripts/ci/check-windows-timestamp-url.mjs

# Ensures Windows installers support manual rollback (downgrades) from the Releases page.
node scripts/ci/check-windows-allow-downgrades.mjs

# Ensures the Tauri updater signing secrets are present for tagged releases.
# (CI reads these from GitHub Actions secrets; locally requires env vars to be set.)
TAURI_PRIVATE_KEY=... TAURI_KEY_PASSWORD=... node scripts/ci/check-tauri-updater-secrets.mjs
# If your private key is unencrypted, the password can be empty/unset:
TAURI_PRIVATE_KEY=... node scripts/ci/check-tauri-updater-secrets.mjs

# Ensures macOS hardened-runtime entitlements include the WKWebView JIT keys
# required for JavaScript/WebAssembly to run in signed/notarized builds.
node scripts/check-macos-entitlements.mjs
```

After all platform builds finish, CI also verifies the **uploaded GitHub Release assets** are
complete and consistent with the Tauri updater manifest (`latest.json`). This prevents publishing a
release where `latest.json` points at missing artifacts or missing signature files.

CI runs:

```bash
node scripts/verify-tauri-latest-json.mjs vX.Y.Z
node scripts/verify-tauri-updater-assets.mjs vX.Y.Z
```

You can run the same checks locally (requires a GitHub token with access to the release assets):

```bash
GITHUB_REPOSITORY=owner/repo GH_TOKEN=... \
  node scripts/verify-tauri-latest-json.mjs vX.Y.Z

GITHUB_REPOSITORY=owner/repo GH_TOKEN=... \
  node scripts/verify-tauri-updater-assets.mjs vX.Y.Z
```

If you already downloaded the manifest files (no GitHub API access needed), you can validate the
updater manifest signature offline with:

```bash
node scripts/ci/verify-updater-manifest-signature.mjs latest.json latest.json.sig
```

To inspect the required platform keys in the manifest:

```bash
jq '.platforms | keys' latest.json
```

Note: `scripts/verify-tauri-latest-json.mjs` delegates to the lower-level validator
`scripts/ci/validate-updater-manifest.mjs` when run in `<tag>` mode, which downloads
`latest.json` / `latest.json.sig` from the draft release and checks targets, signatures, and
referenced assets. The `--manifest/--sig` mode is a lightweight offline check (platform keys +
presence of the signature file).

## Updater restart semantics (important)

When an update is downloaded/installed, the desktop app should restart/exit using Tauri's supported
APIs so the updater plugin can complete any pending work during shutdown.

- Use the backend command **`restart_app`** for updater-driven restarts (it calls
  `AppHandle::restart()` and falls back to `AppHandle::exit(0)`).
- Do **not** use **`quit_app`** for updates. `quit_app` intentionally uses `std::process::exit(0)`
  to avoid re-entering the `CloseRequested` hide-to-tray flow, but hard exits can bypass normal
  shutdown hooks.

## 1) Versioning + tagging

1. Update the desktop app version in `apps/desktop/src-tauri/tauri.conf.json` (`version`).
2. Merge the version bump to `main`.
3. Create and push a tag **with the same version** (CI enforces that the git tag matches
   `tauri.conf.json`):

   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```

The workflow in `.github/workflows/release.yml` will run and create/update a **draft** release with
all platform artifacts attached.

Note: the workflow intentionally pins `tauri-apps/tauri-action` to a specific `v0.x.y` tag to avoid
implicit breakage when the floating `@v0` tag advances. To upgrade, bump the pinned version in
`.github/workflows/release.yml` (and verify a tagged build in CI).

CI note: the workflow also validates that the uploaded Tauri updater manifest (`latest.json` +
`latest.json.sig`) contains entries for every expected OS/architecture target. This catches a common
parallel-build failure mode where the last finishing job overwrites `latest.json` and ships an
incomplete updater manifest.

See `docs/desktop-updater-target-mapping.md` for the exact `latest.json.platforms` key names CI
expects.

## 2) Tauri updater keys (required for auto-update)

Tauri's updater verifies update artifacts using an Ed25519 signature.

This repo already includes the updater **public key** in `apps/desktop/src-tauri/tauri.conf.json`
(`plugins.updater.pubkey`). Tagged releases must be signed with the matching **private key**.

### Store the private key in GitHub Actions

Add the following repository secrets (required for updater signing; password optional):

- `TAURI_PRIVATE_KEY` – the private key string printed by `cargo tauri signer generate` (minisign secret key; base64)
- `TAURI_KEY_PASSWORD` – optional; only needed if the private key was generated with a password

The release workflow passes these to `tauri-apps/tauri-action`, which signs the update artifacts.

CI note:

- In the upstream repo (`wilson-anysphere/formula`), tagged releases (and `workflow_dispatch` runs
  with `upload=true`) will **fail fast** if these secrets are missing (the workflow prints a
  “Missing Tauri updater signing secrets” error). Without the private key, the release workflow
  cannot generate the updater signature files (`*.sig`) required for auto-update.
- On forks/dry-runs without these secrets, the workflow can still build and upload artifacts, but
  updater signature/manifest validation jobs are skipped and auto-update will not work until you
  configure your own updater keypair/secrets.

### (Optional) Generate / rotate a keypair

Only do this if you need to rotate keys (e.g. compromised secret). Run this from the repo root
(requires the Tauri CLI, `cargo tauri`, installed from the `tauri-cli` crate). In agent environments,
use the repo cargo wrapper (`scripts/cargo_agent.sh`) instead of bare `cargo`:

```bash
# (Agents) Initialize safe defaults (memory limits, isolated CARGO_HOME, etc.)
source scripts/agent-init.sh

TAURI_CLI_VERSION=2.9.5

# NOTE: Keep this version in sync with `.github/workflows/release.yml` (env.TAURI_CLI_VERSION).
bash scripts/cargo_agent.sh install tauri-cli --version "$TAURI_CLI_VERSION" --locked --force
(cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate)
```

This prints:
- a **public key** (safe to commit; embed it in `tauri.conf.json`)
- a **private key** (store it as a secret)

### Configure the public key in the app

Update `apps/desktop/src-tauri/tauri.conf.json`:

- `plugins.updater.pubkey` → paste the public key (base64 string)
- `plugins.updater.endpoints` → point at your update JSON endpoint(s)

CI note: tagged releases will fail if `plugins.updater.pubkey` or `plugins.updater.endpoints`
still look like placeholder values. Verify locally with:

```bash
node scripts/check-updater-config.mjs
```

Note: the desktop Rust binary is built with the Cargo feature `desktop` (configured in
`build.features` inside `tauri.conf.json`) so that unit tests can run without the system WebView
toolchain.

## 3) Code signing (optional but recommended)

The release workflow is wired for code signing if the following secrets are present.

CI behavior note:

- If the signing/notarization secrets are **not** configured (common on forks or dry-runs), the
  release workflow will still build successfully and publish **unsigned** artifacts.
- This is implemented by a small CI helper (`scripts/ci/prepare-tauri-signing-config.mjs`) which
  disables the relevant Tauri signing config for that job and only enables notarization when all
  required credentials are available.
- To **enforce** code signing on tagged releases (fail fast if secrets are missing), set the GitHub
  Actions variable `FORMULA_REQUIRE_CODESIGN=1` (Settings → Secrets and variables → Actions →
  Variables). In enforcement mode, CI fails with a “Code signing is required (…) but secrets are
  missing” error and lists the missing secrets.

### macOS (Developer ID + notarization)

Note: `apps/desktop/src-tauri/tauri.conf.json` intentionally does **not** hardcode
`bundle.macOS.signingIdentity`. This keeps local `tauri build` working on macOS machines without
Developer ID certificates installed, and ensures CI uses the explicit `APPLE_SIGNING_IDENTITY`
provided as a secret (avoids ambiguous identity selection when multiple certs exist).

Secrets used by `tauri-apps/tauri-action`:

- `APPLE_CERTIFICATE` – base64-encoded `.p12` Developer ID certificate
- `APPLE_CERTIFICATE_PASSWORD`
- `APPLE_SIGNING_IDENTITY` – optional but recommended; e.g. `Developer ID Application: Your Company (TEAMID)`
- `APPLE_ID` – Apple ID email
- `APPLE_PASSWORD` – app-specific password
- `APPLE_TEAM_ID`

CI preflight: the release workflow validates that `APPLE_CERTIFICATE` is valid base64 and a valid
PKCS#12 archive decryptable with `APPLE_CERTIFICATE_PASSWORD` (fail-fast on tagged releases when
misconfigured). You can run the same check locally:

```bash
APPLE_CERTIFICATE=... APPLE_CERTIFICATE_PASSWORD=... \
  bash scripts/ci/verify-codesign-secrets.sh macos
```

CI guardrail (tagged releases when secrets are configured):

- The release workflow validates that the produced macOS artifacts are **notarized + stapled** so they pass Gatekeeper:
  - `xcrun stapler validate` (requires a stapled notarization ticket)
  - `spctl --assess` (Gatekeeper evaluation)

#### Hardened runtime entitlements (WKWebView / WASM)

The macOS app is signed with the **hardened runtime**. WKWebView (Tauri/Wry) needs explicit JIT entitlements in the signed binary so that JavaScript and WebAssembly can execute reliably.

The entitlements file used during signing is:

- `apps/desktop/src-tauri/entitlements.plist` (wired via `bundle.macOS.entitlements` in `apps/desktop/src-tauri/tauri.conf.json`)

CI guardrail (run on macOS release builds):

```bash
node scripts/check-macos-entitlements.mjs
```

Release workflow note: when macOS signing secrets are configured, CI also runs `codesign -d --entitlements :-` on the built `.app` to ensure these entitlements are actually embedded in the signed bundle (protects against config drift where the plist exists but isn’t used during signing).

If these entitlements are missing, a notarized build can still pass notarization but launch with a **blank window** or a crashing WebView process.

Note: we intentionally do **not** enable `com.apple.security.cs.disable-library-validation` since the app does not load third-party / unsigned dylibs at runtime.

#### Local verification checklist (signed app)

Note: `apps/desktop/src-tauri/tauri.conf.json` does **not** hardcode a signing identity, so a plain
local `tauri build` will typically produce **unsigned** artifacts. To run this checklist locally,
you must build a **signed** app (e.g. use the CI-produced artifacts, or temporarily set
`bundle.macOS.signingIdentity` to your explicit `Developer ID Application: … (TEAMID)` identity and
then revert the change—do not commit it).

1. Build the production bundles:

   ```bash
   pnpm install
   pnpm build:desktop
   cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build
   ```

2. Locate the `.app` produced by Tauri (path can vary by target):

   ```bash
   find apps/desktop/src-tauri/target -maxdepth 8 -type d -path "*/release/bundle/macos/*.app"
   ```

3. Verify the signature + entitlements (replace the path as needed):

   ```bash
   app="$(find apps/desktop/src-tauri/target -maxdepth 8 -type d -path '*/release/bundle/macos/*.app' | head -n 1)"
   echo "Checking app at: $app"
   codesign --verify --deep --strict --verbose=2 "$app"
   codesign -d --entitlements :- "$app" 2>&1 | grep -E "allow-jit|allow-unsigned-executable-memory"
   spctl --assess --type execute -vv "$app"
   ```

4. Launch the app and sanity-check runtime behavior:
   - The window should render (no blank WebView).
   - Network features work (e.g. updater check / HTTPS fetches).
   - Cross-origin isolation still works in the packaged app (see `pnpm -C apps/desktop check:coi`).

#### Troubleshooting: blank window / crashes in a signed build

If a signed/notarized build launches with a blank window or crashes immediately, check:

1. The **entitlements actually embedded in the signed app** (not just the plist file in the repo):

   ```bash
   codesign -d --entitlements :- /Applications/Formula.app
   ```

   Ensure it includes:
   - `com.apple.security.cs.allow-jit`
   - `com.apple.security.cs.allow-unsigned-executable-memory`

2. The signature is valid and the hardened runtime is enabled:

   ```bash
   codesign -dv --verbose=4 /Applications/Formula.app 2>&1 | grep -E "Runtime|TeamIdentifier|Identifier"
   ```

3. macOS logs/crash reports:
   - Use **Console.app** → Crash Reports / log stream.
   - Look for `WebKit`, `JavaScriptCore`, or `EXC_BAD_ACCESS` crashes in a `WebContent` process.
### Windows (Authenticode)

Secrets:

- `WINDOWS_CERTIFICATE` – base64-encoded `.pfx`
- `WINDOWS_CERTIFICATE_PASSWORD`

CI preflight: the release workflow validates that `WINDOWS_CERTIFICATE` is valid base64 and a valid
PKCS#12 archive decryptable with `WINDOWS_CERTIFICATE_PASSWORD` (fail-fast when misconfigured). You
can run the same check locally:

```bash
WINDOWS_CERTIFICATE=... WINDOWS_CERTIFICATE_PASSWORD=... \
  bash scripts/ci/verify-codesign-secrets.sh windows
```

Verification (signed artifacts):

- After a Windows release build, verify the generated installer(s) are Authenticode-signed and timestamped.
  Note: bundle output paths can vary depending on whether you built with an explicit `--target <triple>`
  (CI does) — adjust the `target\\...` path accordingly.
  ```powershell
  signtool verify /pa /v apps\desktop\src-tauri\target\release\bundle\nsis\*.exe
  signtool verify /pa /v apps\desktop\src-tauri\target\release\bundle\msi\*.msi
  ```
  - Release CI also runs `signtool verify /pa /v` on the produced installers when `WINDOWS_CERTIFICATE`
    is configured (see `scripts/ci/check-windows-installer-signatures.py`).

Timestamping:

- The Authenticode timestamp server is configured in `apps/desktop/src-tauri/tauri.conf.json` under
  `bundle.windows.timestampUrl` (currently `https://timestamp.digicert.com`).
  - Release CI preflight enforces this uses HTTPS (see `scripts/ci/check-windows-timestamp-url.mjs`).
- If a release fails due to timestamping/network issues, switch `timestampUrl` to another **HTTPS**
  timestamp server provided/recommended by your signing certificate vendor and re-run the workflow.

## Windows ARM64 build prerequisites (MSVC)

The release workflow builds **Windows ARM64** installers by cross-compiling from an x64 GitHub-hosted
Windows runner (MSVC target `aarch64-pc-windows-msvc`).

This requires the Visual Studio **MSVC ARM64** toolchain to be present on the runner:

- Visual Studio component: `Microsoft.VisualStudio.Component.VC.Tools.ARM64`

You also need the Rust standard library for the ARM64 target installed (CI does this automatically
in the release workflow):

```bash
rustup target add aarch64-pc-windows-msvc
```

Windows ARM64 builds also require a Windows SDK installation with **ARM64** libraries present
(UM + UCRT). CI sanity-checks for directories like:

- `C:\Program Files (x86)\Windows Kits\10\Lib\<version>\um\arm64`
- `C:\Program Files (x86)\Windows Kits\10\Lib\<version>\ucrt\arm64`

When cross-compiling locally from an x64 Windows machine, run the build in a Visual Studio
Developer Prompt configured for **amd64 → arm64** (CI uses `ilammy/msvc-dev-cmd` with `arch:
amd64_arm64`).

Then build the ARM64 installers:

```powershell
cd apps/desktop
cargo tauri build --target aarch64-pc-windows-msvc --bundles msi,nsis
```

Expected outputs:

- `apps/desktop/src-tauri/target/aarch64-pc-windows-msvc/release/bundle/msi/*.msi`
- `apps/desktop/src-tauri/target/aarch64-pc-windows-msvc/release/bundle/nsis/*.exe`

Sanity-check (optional): verify the built desktop executable is actually ARM64 (AA64):

```powershell
dumpbin /headers apps/desktop\src-tauri\target\aarch64-pc-windows-msvc\release\formula-desktop.exe `
  | Select-String -Pattern 'machine' -CaseSensitive:$false
```

Expected output includes `machine (AA64)`.

GitHub-hosted runner images do not always include this workload by default. The release workflow
checks for a complete ARM64 MSVC + SDK toolchain:

- MSVC: `VC\\Tools\\MSVC\\*\\lib\\arm64` + `VC\\Tools\\MSVC\\*\\bin\\Hostx64\\arm64\\{cl.exe,link.exe}`
- Windows SDK: `Windows Kits\\10\\Lib\\*\\{um,ucrt}\\arm64`

If any of these are missing, CI installs the MSVC ARM64 component via `vs_installer.exe` and fails
with a clear error if the runner image still lacks the required ARM64 SDK libraries.

CI smoke test:

- `.github/workflows/windows-arm64-smoke.yml` runs `cargo tauri build --target aarch64-pc-windows-msvc`
  and asserts that the expected Windows bundles land under
  `apps/desktop/src-tauri/target/aarch64-pc-windows-msvc/release/bundle/**`.

## Windows installer bundler prerequisites (WiX + NSIS)

Formula ships **both** Windows installer formats for **x64** and **ARM64**:

- **MSI** (WiX Toolset; Tauri uses `candle.exe` + `light.exe`)
- **EXE** (NSIS; Tauri uses `makensis.exe`)

In CI, `.github/workflows/release.yml` installs these tools automatically via Chocolatey so tagged
releases always include both `.msi` and `.exe` assets for each architecture.

For local Windows builds, ensure WiX + NSIS are installed and on `PATH` (example using Chocolatey):

```powershell
choco install wixtoolset nsis --yes --no-progress
```

## Windows: WebView2 runtime installation (required)

Formula relies on the **Microsoft Edge WebView2 Evergreen Runtime** on Windows. The Windows installers are configured to
install WebView2 automatically if it is missing by using Tauri's WebView2 installer integration:

- Config: `apps/desktop/src-tauri/tauri.conf.json` → `bundle.windows.webviewInstallMode.type = "downloadBootstrapper"`
  (Evergreen **bootstrapper**; we also set `silent: true`; requires an internet connection to download/install the runtime).
- Alternative: `bundle.windows.webviewInstallMode.type = "embedBootstrapper"` bundles the bootstrapper into the installer
  (~1.8 MB larger installer; still requires internet to install the runtime; can be more reliable on older Windows versions).
- Note: Tauri defaults to `downloadBootstrapper` if `webviewInstallMode` is omitted, but we keep it **explicit** in config
  (and guardrailed in CI/tests) so the behavior is obvious and doesn't regress silently.
- CI verification:
  - Fast preflight (config-only): `node scripts/ci/check-webview2-install-mode.mjs`
  - Built-artifact inspection: `python scripts/ci/check-windows-webview2-installer.py` (asserts the produced installers
    contain a WebView2 bootstrapper/runtime reference), failing the release build if this regresses.
    - Note: this inspection also supports the `fixedRuntime` mode (it will look for fixed runtime payload files).

To verify locally after a Windows build, run:

```bash
node scripts/ci/check-webview2-install-mode.mjs
python scripts/ci/check-windows-webview2-installer.py
```

To verify on a clean Windows VM (no preinstalled WebView2 Runtime):

1. Ensure **Microsoft Edge WebView2 Runtime** is not installed (Windows Settings → Apps).
2. Run the Formula installer.
3. Confirm the app launches successfully (the installer should bootstrap WebView2 automatically).

If you need an offline-friendly installer, change `bundle.windows.webviewInstallMode` to `offlineInstaller` or
`fixedRuntime` (at the cost of a much larger installer: roughly **+127 MB** for `offlineInstaller` or **+180 MB** for
`fixedRuntime`).

For the full set of options, see the Tauri docs:
https://v2.tauri.app/distribute/windows-installer/#webview2-installation-options

If you need to point end users at a manual install (e.g. no internet / locked-down environments),
Microsoft’s official WebView2 download page is:
https://developer.microsoft.com/en-us/microsoft-edge/webview2/

Troubleshooting note: WebView2 can be installed per-user or per-machine. If a user reports that
Formula “can’t find WebView2” even though it is installed, ensure the runtime is installed for the
same Windows user account (or install it system-wide).

## 4) Hosting updater endpoints

The desktop app is configured to use **GitHub Releases** as the updater source.

`apps/desktop/src-tauri/tauri.conf.json` points at the `latest.json` manifest generated and uploaded
by `tauri-apps/tauri-action@v0.6.1`:

```
https://github.com/wilson-anysphere/formula/releases/latest/download/latest.json
```

To sanity-check your updater configuration before tagging a release, run:

```bash
node scripts/check-updater-config.mjs
```

Notes:

- The GitHub `/releases/latest` URL only tracks the **latest published** (non-draft) release.
  Draft releases created by the workflow are for QA and will not be picked up by auto-update until
  you click **Publish release**.
- `tauri-action` also uploads a corresponding signature file (`latest.json.sig`), which the updater
  verifies using the `plugins.updater.pubkey` embedded in the app.

If you fork this repo, change the endpoint to match your GitHub org/repo.

### Optional: custom update server / CDN

If you don't want clients to fetch update metadata from GitHub directly, you can mirror the release
assets (including `latest.json` + `latest.json.sig`) to your own host and update
`plugins.updater.endpoints` accordingly.

### Updater targets (`{{target}}`) and `latest.json`

Tauri’s updater chooses which file to download by matching the running app’s **target string**
(`{{target}}` in templated endpoints) against the keys under `platforms` in `latest.json`.

When using GitHub Releases, `tauri-action` generates `latest.json` and uploads it to the release
alongside the installers. The file is structured roughly like:

```jsonc
{
  "version": "0.1.0",
  "notes": "...",
  "pub_date": "2026-01-01T00:00:00Z",
  "platforms": {
    "darwin-universal": { "url": "…", "signature": "…" },
    "windows-x86_64": { "url": "…", "signature": "…" },
    "windows-aarch64": { "url": "…", "signature": "…" },
    "linux-x86_64": { "url": "…", "signature": "…" }
  }
}
```

Expected `{{target}}` values for this repo’s release matrix:

- **macOS (universal):** `darwin-universal` (some toolchains use `universal-apple-darwin`) pointing
  at the updater payload, typically an `.app.tar.gz`.
- **Windows:** `windows-x86_64` (some toolchains use `x86_64-pc-windows-msvc`) and
  `windows-aarch64` / `windows-arm64` (some toolchains use `aarch64-pc-windows-msvc`) — one entry per
  architecture. Each entry points at the **updater-consumed** Windows installer (see
  `docs/desktop-updater-target-mapping.md`; currently the updater entries should point at the
  **`.msi`** for both x64 and ARM64).
- **Linux:** `linux-x86_64` (some toolchains use `x86_64-unknown-linux-gnu`; points at the updater
  payload, typically the `.AppImage`).

Note: `apps/desktop/src-tauri/tauri.conf.json` sets `bundle.targets: "all"`, which enables all
supported bundlers for the current platform (including **MSI/WiX** + **NSIS** on Windows). CI still
passes `--bundles msi,nsis` and installs WiX + NSIS explicitly so Windows releases always include
both installer formats.

For reference, this is how the release workflow’s Tauri build targets map to updater targets:

| Workflow build | Tauri build args | Rust target triple | `latest.json` platform key(s) |
| --- | --- | --- | --- |
| macOS universal | `--target universal-apple-darwin` | `aarch64-apple-darwin` + `x86_64-apple-darwin` | `darwin-universal` (or `universal-apple-darwin`) |
| Windows x64 | `--target x86_64-pc-windows-msvc --bundles msi,nsis` | `x86_64-pc-windows-msvc` | `windows-x86_64` (or `x86_64-pc-windows-msvc`) |
| Windows ARM64 | `--target aarch64-pc-windows-msvc --bundles msi,nsis` | `aarch64-pc-windows-msvc` | `windows-aarch64` / `windows-arm64` (or `aarch64-pc-windows-msvc`) |
| Linux x64 | `--bundles appimage,deb,rpm` | `x86_64-unknown-linux-gnu` | `linux-x86_64` (or `x86_64-unknown-linux-gnu`) |

Note: `.deb` and `.rpm` are shipped for manual install/downgrade, but are not typically used by the
Tauri updater on Linux. If a target entry is missing from `latest.json`, auto-update for that
platform/arch will not work even if the GitHub Release has other assets attached.

## Linux: compatibility expectations (`.AppImage` vs `.deb`/`.rpm`)

The Linux desktop shell uses the system WebView provided by **WebKitGTK** (Tauri/Wry). This repo
targets **WebKitGTK 4.1 + GTK3**, so the distro-native packages (`.deb` / `.rpm`) are most
compatible with distros that ship those versions.

- Prefer **`.deb` / `.rpm`** when the target distro provides WebKitGTK 4.1 (Debian/Ubuntu:
  `libwebkit2gtk-4.1-0`; Fedora: `webkit2gtk4.1`). These integrate with the system package manager
  and are the expected “happy path” on modern distros.
  - Note: some RHEL 9-family distros ship `webkit2gtk3` (WebKitGTK 4.0) instead of WebKitGTK 4.1;
    in that case prefer the `.AppImage`.
- Prefer the **`.AppImage`** when installing the `.deb`/`.rpm` fails due to missing or incompatible
  system libraries (commonly WebKitGTK/GTK3). The AppImage bundles more runtime libraries and tends
  to run on a wider range of distros.

### Quick compatibility check

On the target distro, confirm a WebKitGTK 4.1 runtime package is available via the package manager:

- Debian/Ubuntu: `apt-cache policy libwebkit2gtk-4.1-0` (or `apt search libwebkit2gtk-4.1`)
- Fedora: `dnf info webkit2gtk4.1`
- RHEL 9-family: `dnf info webkit2gtk3` (WebKitGTK 4.0; expect to use the `.AppImage` if 4.1 is unavailable)

If the distro cannot install a WebKitGTK 4.1 package, recommend the `.AppImage` instead of the
`.deb`/`.rpm`.

## Linux: `.deb` runtime dependencies (WebView + tray)

The Linux desktop shell is built on **GTK3 + WebKitGTK** (Tauri/Wry) and uses the **AppIndicator**
stack for the tray icon.

The resulting `.deb` must declare the required runtime packages so that:

- the WebView can start (no missing `libwebkit2gtk*` / `libgtk*` shared libraries)
- the tray icon backend can be loaded (otherwise the tray icon will be missing)

These dependencies are declared in `apps/desktop/src-tauri/tauri.conf.json` under
`bundle.linux.deb.depends`:

- `libwebkit2gtk-4.1-0` – WebKitGTK system WebView used by Tauri on Linux.
- `libgtk-3-0t64 | libgtk-3-0` – GTK3 (windowing/event loop; also required by WebKitGTK).
  Ubuntu 24.04 uses `*t64` package names for some libraries due to the `time_t` 64-bit transition.
- `libayatana-appindicator3-1 | libappindicator3-1` – tray icon backend.
  The Rust bindings (`libappindicator-sys`) load this library dynamically at runtime; without it
  the app can launch but the tray icon will not appear.
- `librsvg2-2` – SVG rendering used by parts of the GTK icon stack / common icon themes.
- `libssl3t64 | libssl3` – OpenSSL runtime required by native dependencies in the Tauri stack
  (Ubuntu 24.04 uses the `libssl3t64` package name).

### Validating the `.deb`

After building via `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` (or after CI produces an artifact),
verify the dependency list and shared library resolution.

From `apps/desktop/src-tauri`:

```bash
# Inspect the control file (check Depends: ...)
deb="$(ls target/release/bundle/deb/*.deb | head -n 1)"
dpkg -I "$deb"

# Extract and confirm all linked shared libraries resolve
tmpdir="$(mktemp -d)"
dpkg-deb -x "$deb" "$tmpdir"
ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
```

For a clean install test (no GUI required), use a container:

```bash
docker run --rm -it \
  -v "$PWD/target/release/bundle/deb:/deb" \
  ubuntu:24.04 bash -lc '
    apt-get update
    apt-get install -y --no-install-recommends /deb/*.deb
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

CI guardrails (tagged releases):

- `bash scripts/ci/verify-linux-package-deps.sh` inspects the produced `.deb` with `dpkg -I` / `dpkg-deb -f` and fails the
  workflow if the **core runtime dependencies** are missing from `Depends:`.
- `bash scripts/ci/linux-package-install-smoke.sh deb` installs the `.deb` into a clean Ubuntu container and fails if
  `ldd /usr/bin/formula-desktop` reports missing shared libraries.

## Linux: `.rpm` runtime dependencies (Fedora/RHEL-family + openSUSE)

For RPM-based distros (Fedora/RHEL/CentOS derivatives), the same GTK3/WebKitGTK/AppIndicator stack
must be present at runtime.

These dependencies are declared in `apps/desktop/src-tauri/tauri.conf.json` under
`bundle.linux.rpm.depends`.

We use **RPM rich dependencies** (`(a or b)`) so the RPM declares correct runtime requirements across
common RPM families (Fedora/RHEL vs openSUSE), which sometimes use different package names for the
same shared libraries.

Effective package names (varies by distro):

- WebKitGTK 4.1 runtime:
  - Fedora/RHEL: `webkit2gtk4.1`
  - openSUSE: `libwebkit2gtk-4_1-0`
  - Note: some RHEL-family distros ship WebKitGTK 4.0 as `webkit2gtk3` and may not be compatible
    with this build.
- GTK3 runtime:
  - Fedora/RHEL: `gtk3`
  - openSUSE: `libgtk-3-0`
- AppIndicator/Ayatana tray backend:
  - Fedora/RHEL: `libayatana-appindicator-gtk3` or `libappindicator-gtk3`
  - openSUSE: `libayatana-appindicator3-1` or `libappindicator3-1`
- librsvg runtime:
  - Fedora/RHEL: `librsvg2`
  - openSUSE: `librsvg-2-2`
- OpenSSL runtime:
  - Fedora/RHEL: `openssl-libs`
  - openSUSE: `libopenssl3`

Note: the AppIndicator dependency is expressed using RPM “rich dependency” syntax (`(A or B)`).
This requires a modern RPM stack (rpm ≥ 4.12). On older RPM-based distros, prefer the `.AppImage`.

### Validating the `.rpm`

After building via `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` (or after CI produces an artifact),
verify the `Requires:` list and shared library resolution.

From `apps/desktop/src-tauri`:

```bash
# Inspect declared dependencies (check webkit2gtk/gtk3/appindicator/etc)
rpm_pkg="$(ls target/release/bundle/rpm/*.rpm | head -n 1)"
rpm -qpR "$rpm_pkg"

# Extract and confirm all linked shared libraries resolve
# (requires `cpio`: Fedora `dnf -y install cpio`, Debian/Ubuntu `apt-get install -y cpio`)
tmpdir="$(mktemp -d)"
(cd "$tmpdir" && rpm2cpio "$rpm_pkg" | cpio -idmv)
ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
```

For a clean install test (no GUI required), use a Fedora container:

```bash
docker run --rm -it \
  -v "$PWD/target/release/bundle/rpm:/rpm" \
  fedora:40 bash -lc '
    # The Tauri updater `.sig` files are *not* RPM GPG signatures, so install with --nogpgcheck.
    dnf -y install --nogpgcheck --setopt=install_weak_deps=False /rpm/*.rpm
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

Optional: openSUSE smoke install (helps validate our RPM rich-deps cover openSUSE package naming):

```bash
docker run --rm -it \
  -v "$PWD/target/release/bundle/rpm:/rpm" \
  opensuse/tumbleweed:latest bash -lc '
    zypper --non-interactive refresh
    zypper --non-interactive install --no-recommends --allow-unsigned-rpm /rpm/*.rpm
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

CI guardrails (tagged releases):

- `bash scripts/ci/verify-linux-package-deps.sh` inspects the produced `.rpm` with `rpm -qpR` and fails the workflow if the
  **core runtime dependencies** are missing from the RPM metadata.
- `bash scripts/ci/linux-package-install-smoke.sh rpm` installs the `.rpm` into a clean Fedora container and fails if
  `ldd /usr/bin/formula-desktop` reports missing shared libraries.

Note: showing a tray icon also requires a desktop environment with **StatusNotifier/AppIndicator**
support (e.g. the GNOME Shell “AppIndicator and KStatusNotifierItem Support” extension).

## 5) Verifying a release

After the workflow completes, open the GitHub Release (draft) and confirm the expected artifacts
are attached:

1. Open the GitHub Release (draft) and confirm:
   - Updater metadata: `latest.json` and `latest.json.sig`
   - `SHA256SUMS.txt` (SHA256 checksums for all release assets)
   - macOS (**universal**): `.dmg` (installer) + `.app.tar.gz` (updater payload)
   - Windows **x64**: installers (WiX `.msi` **and** NSIS `.exe`, filename typically includes `x64` / `x86_64`)
   - Windows **ARM64**: installers (WiX `.msi` **and** NSIS `.exe`, filename typically includes `arm64` / `aarch64`)
   - Linux: `.AppImage` + `.deb` + `.rpm`

   This repo requires Tauri updater signing for tagged releases, so expect `.sig` signature files to
   be uploaded alongside the produced artifacts:
   - macOS: `.dmg.sig` and `.app.tar.gz.sig`
   - Windows (each architecture): `.msi.sig` and `.exe.sig`
   - Linux: `.AppImage.sig`, `.deb.sig`, `.rpm.sig`

   (These `.sig` files are Tauri/Ed25519 updater signatures, **not** OS/package-manager signatures.)

   Note: the release workflow enforces that each Windows target produces **both** a `.msi` and a
   `.exe` installer under `apps/desktop/src-tauri/target/<triple>/release/bundle/**`. If the MSI
   bundler regresses (e.g. missing WiX toolset support for ARM64), the workflow fails so we don’t
   ship a Windows release that violates the distribution requirement.

   Note: even though the Tauri updater typically uses the `.AppImage` on Linux, we still ship
   distro-native packages (`.deb`/`.rpm`) for manual install/downgrade and their corresponding `.sig`
   files.

   If an expected platform/arch is missing entirely, start with the GitHub Actions run for that tag
   and check the build job for the relevant platform/target (and whether the Tauri bundler step
   failed before uploading assets).
2. Download `latest.json` and confirm `platforms` includes entries for:
   - `darwin-universal`
   - `windows-x86_64`
   - `windows-aarch64`
   - `linux-x86_64`

   Note: the tagged-release CI validator is intentionally **strict** about these key names. If you
   see different spellings (for example Rust target triples like `x86_64-pc-windows-msvc` or
   `aarch64-pc-windows-msvc`), that usually indicates a Tauri/toolchain change and the
   docs/validator should be updated together.

   Quick check (after downloading `latest.json` to your current directory):

   ```bash
   python - <<'PY'
   import json
   data = json.load(open("latest.json", encoding="utf-8"))
   keys = sorted((data.get("platforms") or {}).keys())
   print("\n".join(keys) if keys else "(no platforms found)")
   PY
   ```

   Also confirm each platform entry points at the **updater-consumed** asset type:
   - `darwin-*` → `*.app.tar.gz`
   - `windows-*` → `*.msi` (updater runs the Windows Installer)
   - `linux-*` → `*.AppImage`

3. Download the artifacts and do quick sanity checks:

   ### macOS: confirm the app is universal

   Run `lipo -info` on the bundled executable (`Formula.app/Contents/MacOS/formula-desktop`):

   ```bash
   # Option A: from the .app.tar.gz
   app_tgz="$(ls *.app.tar.gz | head -n 1)"
   tar -xzf "$app_tgz"
   lipo -info "Formula.app/Contents/MacOS/formula-desktop"

   # Expected output includes both: x86_64 arm64
   ```

   If you only have a `.dmg`, mount it and inspect the `.app` inside:

   ```bash
   dmg="$(ls *.dmg | head -n 1)"
   mnt="$(mktemp -d)"
   hdiutil attach "$dmg" -nobrowse -mountpoint "$mnt"
   lipo -info "$mnt/Formula.app/Contents/MacOS/formula-desktop"
   hdiutil detach "$mnt"
   ```

   ### Windows: confirm x64 vs arm64 + signatures (if enabled)

   Check the machine type of the installed/bundled executable using `dumpbin` (Visual Studio tools):

   ```bat
   REM From a "Developer Command Prompt for VS"
   dumpbin /headers path\to\formula-desktop.exe | findstr /i machine

   REM x64  => machine (8664)
   REM arm64 => machine (AA64)
   ```

   If Authenticode signing is enabled, verify signatures:

   ```bat
   signtool verify /pa /v path\to\installer.exe
   signtool verify /pa /v path\to\installer.msi
   ```

   ### Windows: WebView2 install smoke test (clean VM)

   On a clean Windows VM **without** WebView2 (or after uninstalling **Microsoft Edge WebView2 Runtime**),
   run the installer. It should install WebView2 via the configured Evergreen bootstrapper and then
   the app should launch successfully. (This requires an internet connection when using the bootstrapper modes.)

   ### Linux: inspect dependencies + `ldd` smoke check

   ```bash
   # Dependency metadata (ensure the runtime deps are present)
   deb="$(ls *.deb | head -n 1)"
   rpm="$(ls *.rpm | head -n 1)"

   dpkg -I "$deb"
   rpm -qpR "$rpm"

   # Extract and ensure the main binary has no missing shared libraries.
   tmpdir="$(mktemp -d)"
   dpkg-deb -x "$deb" "$tmpdir"
   ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
   ```

4. Download/install on each platform (matching the architecture).
5. Publish the release to make it visible to users and (if your updater endpoint references
   GitHub) available for auto-update.

### Verifying installer checksums

Each tagged release includes a `SHA256SUMS.txt` asset. To verify a download:

1. Download the installer/bundle you want **and** `SHA256SUMS.txt` from the same release.
2. Compute the SHA256 locally and compare it to the matching line in `SHA256SUMS.txt`:

   ```bash
   # macOS
   shasum -a 256 Formula*.dmg

   # Linux
   sha256sum Formula*.AppImage
   ```

   ```powershell
   # Windows (PowerShell)
   Get-FileHash -Algorithm SHA256 .\Formula*.msi
   ```

Also verify **cross-origin isolation** is enabled in the packaged app (required for `SharedArrayBuffer` and the Pyodide Worker backend):

- From source (recommended preflight): `pnpm -C apps/desktop check:coi`
- Or in an installed build: ensure there is no startup toast complaining about missing cross-origin isolation, and (if you have DevTools)
   confirm `globalThis.crossOriginIsolated === true`.

### File associations + deep link scheme (CI guardrailed)

The release workflow also inspects the **built artifacts** to ensure OS
integration metadata made it into the final bundles (not just `tauri.conf.json`):

- macOS: `CFBundleDocumentTypes` includes `.xlsx`/`.csv`/`.parquet` (etc) and
  `CFBundleURLTypes` includes the `formula` scheme.
- Linux: the installed `.desktop` file advertises the expected `MimeType=` list
  and has an `Exec=` placeholder so double-click open passes a path/URL.

You can run the same checks locally after building:

```bash
# macOS
app="$(find apps/desktop/src-tauri/target/release/bundle/macos -maxdepth 2 -name '*.app' -print -quit)"
plutil -p "$app/Contents/Info.plist" | head -n 200
python scripts/ci/verify_macos_bundle_associations.py --info-plist "$app/Contents/Info.plist"

# Linux (.deb)
deb="$(find apps/desktop/src-tauri/target/release/bundle/deb -maxdepth 1 -name '*.deb' -print -quit)"
tmpdir="$(mktemp -d)"
dpkg-deb -x "$deb" "$tmpdir"
python scripts/ci/verify_linux_desktop_integration.py --deb-root "$tmpdir"
```

CI note: tagged releases run this check on macOS/Windows/Linux before uploading artifacts. If you need to temporarily skip the
check on macOS/Windows (e.g. a hosted-runner regression makes it flaky), set the GitHub Actions variable
`FORMULA_COI_CHECK_ALL_PLATFORMS=0` (or `false`) to keep the Linux check while disabling the non-Linux ones.

## 6) Installer/bundle artifact size reporting + size gate (tagged releases enforced)

The release workflow reports the size of each generated installer/bundle artifact (DMG / MSI / EXE /
AppImage / DEB / RPM / etc) in the GitHub Actions **step summary**, and **fails tagged releases** if
any artifact exceeds the per-artifact size budget (default: **50 MB**).

Note: this is an **installer artifact** budget (DMG/MSI/AppImage/etc), not the **frontend asset
download size** budget (compressed JS/CSS/WASM; see `pnpm -C apps/desktop check:bundle-size` and
`python scripts/desktop_size_report.py`).

### Rust binary size controls (Cargo release profile)

The largest contributor under our control is the Rust desktop binary (`formula-desktop`). Size is
primarily controlled by the workspace Cargo release profile in the repo root `Cargo.toml`:

- `strip = "symbols"` – ensures release binaries do not ship with symbol/debug info.
- `lto = "thin"` – enables ThinLTO (often shrinks binaries and improves runtime perf).
- `codegen-units = 1` – improves LTO effectiveness and typically reduces size.

The release workflow also runs `python scripts/verify_desktop_binary_stripped.py` after building to
fail the workflow if the produced desktop binary is not stripped (or if symbol sidecar files like
`.pdb`/`.dSYM` end up in the bundle output directory).

Local note: `scripts/cargo_agent.sh` sets `CARGO_PROFILE_RELEASE_CODEGEN_UNITS` by default for
stability on multi-agent hosts. If you want local builds to match CI's `codegen-units = 1`, run:

```bash
export CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1
(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)
```

### Configuration / override (GitHub Actions variables)

The tagged release workflow defaults to `FORMULA_ENFORCE_BUNDLE_SIZE=1`. To override, set repository
variables in **Settings → Secrets and variables → Actions → Variables**:

- `FORMULA_BUNDLE_SIZE_LIMIT_MB=50` → override the default **50 MB** per-artifact budget
- `FORMULA_ENFORCE_BUNDLE_SIZE=0` (or `false`) → disable enforcement for exceptional releases

### Run the size check locally

1. Build the desktop bundles for your platform:

    ```bash
    source scripts/agent-init.sh

    TAURI_CLI_VERSION=2.9.5
    bash scripts/cargo_agent.sh install tauri-cli --version "$TAURI_CLI_VERSION" --locked --force
    (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)
    ```

2. From the repo root, print an installer/bundle artifact size table:

   ```bash
   python scripts/desktop_bundle_size_report.py
   ```

   For a machine-readable report (useful for CI debugging):

   ```bash
   python scripts/desktop_bundle_size_report.py --json target/desktop-bundle-size-report.json
   ```

3. (Optional) enforce the budget locally:

    ```bash
    FORMULA_ENFORCE_BUNDLE_SIZE=1 FORMULA_BUNDLE_SIZE_LIMIT_MB=50 \
      python scripts/desktop_bundle_size_report.py
    ```

## 7) Rollback / downgrade support (do not delete old releases)

The platform requirement **"Rollback capability"** is satisfied by supporting a **user-facing
downgrade path**:

- Users can open the Formula GitHub Releases page, download a prior version for their platform, and
  install it (see rollback notes in `docs/11-desktop-shell.md`):
  https://github.com/wilson-anysphere/formula/releases

This only works if older release assets remain available.

### Windows downgrade notes (MSI / EXE)

Windows installers often block installing an older version over a newer one. Formula’s Windows
bundles are explicitly configured to **allow downgrades** via
`bundle.windows.allowDowngrades: true` in `apps/desktop/src-tauri/tauri.conf.json`.

Expected behavior when downgrading **manually** from the GitHub Releases page:

- **NSIS `.exe` installer:** detects the newer installed version and shows a maintenance screen.
  For the cleanest rollback, choose **“Uninstall before installing”**, then proceed with the
  install.
- **WiX `.msi` installer:** if your currently installed Formula version was installed via **MSI**
  (including installs performed by the in-app auto-updater), running an older MSI will remove the
  installed MSI version and then install the selected version (major upgrade flow with downgrades
  enabled).

Tip: prefer using the **same installer format** you originally installed with (`.exe` ↔ `.exe`, or
`.msi` ↔ `.msi`). Switching formats can result in a second installation; if in doubt, uninstall the
current version first.

**Maintainer verification checklist (Windows)**

Before publishing a release, sanity-check the rollback path on a Windows machine/VM:

1. Install a newer build (e.g. `vX.Y.Z`) using either the `.exe` or `.msi`.
2. Run the **older** installer (e.g. `vX.Y.(Z-1)`) of the **same format**:
   - `.exe` downgrade: installer should show the maintenance screen; choose **“Uninstall before installing”**.
   - `.msi` downgrade: installer should proceed and end with the older version installed.
3. Launch the app and confirm the reported version matches the older build.

If an installer refuses to proceed (e.g. “a newer version is already installed”), uninstall the
newer version first from **Settings → Apps → Installed apps**, then install the older `.msi`/`.exe`.

**Release hygiene requirements**

1. **Do not delete prior GitHub Releases or assets.**
   - Keep at least several older versions available so users can downgrade when needed.
2. If you mirror artifacts to `releases.formula.app` (or another CDN), ensure you **retain older
   installers/bundles** there too.
   - Users may need to roll back even if the app can't start, so the download URLs must work
     without relying on the updater UI.
