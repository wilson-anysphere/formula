# Desktop release process

This repository ships the desktop app via GitHub Releases and Tauri's built-in updater.
Tagged pushes (`vX.Y.Z`) trigger a GitHub Actions workflow that builds installers/bundles for
macOS (Apple Silicon + Intel), Windows, and Linux and uploads them to a **draft** GitHub Release.

Platform/architecture expectations for a release:

- **macOS:** **universal** build (Intel + Apple Silicon): `.dmg` (manual installer) + updater tarball
  (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) (auto-update payload).
- **Windows:** **x64** + **ARM64**: `.msi` (manual installer **and** auto-update payload) + `.exe`
  (manual installer) for each architecture.
- **Linux:** **x86_64** + **ARM64**: `.AppImage` (auto-update payload; also a portable/manual option)
  + `.deb` + `.rpm` (manual installers) for each architecture.

The workflow also uploads updater metadata (`latest.json` + `latest.json.sig`) used by the Tauri
updater. Note: `latest.json.sig` signs the manifest itself; each platform’s update payload signature
is carried in `latest.json.platforms[*].signature` and is also uploaded as a detached `<asset>.sig`
file for offline verification.

CI also runs a lightweight **Linux AppImage smoke test** (no GUI) to catch common packaging issues
early:

- Extract the generated `*.AppImage` via `--appimage-extract` (no FUSE required)
- Verify the extracted main ELF binary architecture via `readelf -h`
- Run `ldd` and fail the workflow if any shared libraries are `not found`
- Then validate desktop integration metadata from the extracted payload:
  - file associations from `bundle.fileAssociations` (including Parquet)
  - deep-link scheme handler(s) (`x-scheme-handler/<scheme>`, e.g. `formula://...`)
  - `Exec=` placeholder (`%u/%U/%f/%F`) so file opens + deep links can pass a path/URL
  - OSS/compliance artifacts (`LICENSE`/`NOTICE`) under `/usr/share/doc/<mainBinaryName>/`
  - Parquet shared-mime-info definition packaging (`/usr/share/mime/packages/<identifier>.xml` with a `*.parquet` glob)

See `scripts/ci/check-appimage.sh`.
See also `scripts/validate-linux-appimage.sh` (desktop integration metadata + bundle version checks vs `tauri.conf.json`).

Tip: `scripts/validate-linux-appimage.sh` supports an optional exec sanity check that runs the
extracted `AppRun` entrypoint in a quick-exit mode:

```bash
# Runs squashfs-root/AppRun --startup-bench under a short timeout (uses Xvfb when needed).
bash scripts/validate-linux-appimage.sh --exec-check --exec-timeout 30
```

For Linux DEB bundles, see `scripts/validate-linux-deb.sh` (host `dpkg-deb` metadata checks including version/package name vs `tauri.conf.json` + optional Ubuntu container install smoke test).

For Linux RPM bundles, see `scripts/validate-linux-rpm.sh` (host `rpm -qp` metadata checks including version/package name vs `tauri.conf.json` + optional Fedora container install smoke test).

For the **exact** `latest.json.platforms` key names (and which asset each key should point to),
see:

- `docs/desktop-updater-target-mapping.md`

## Rust toolchain pinning (release stability)

Desktop packaging is sensitive to Rust/toolchain changes. This repo pins Rust via
`rust-toolchain.toml` at the repo root, and CI/release workflows enforce that they install the
same version.

Note: the `RUSTUP_TOOLCHAIN` environment variable has higher precedence than `rust-toolchain.toml`.
If it is set globally (often to `stable`), it can accidentally bypass the pin for local builds and
preflight scripts. Unset it (or `. scripts/agent-init.sh` / use `scripts/cargo_agent.sh`) when
running release steps locally.

To upgrade Rust, open a PR that bumps `rust-toolchain.toml` (and any workflow pins/comments that the
CI guard requests) and rely on CI to validate the new toolchain before tagging a release.

## GitHub Actions runner OS pinning (release stability)

The Desktop Release workflow pins GitHub-hosted runner image versions (for example `macos-14`,
`windows-2022`, `ubuntu-24.04`) instead of using moving `*-latest` labels. GitHub periodically
updates the `*-latest` aliases to newer OS versions, which can unexpectedly break release builds
and make historical tagged builds harder to reproduce.

See https://github.com/actions/runner-images for the currently supported runner images and any
deprecation notices.

To update runner pins, open a PR that bumps the runner labels in
`.github/workflows/release.yml` **after** validating that the release workflow succeeds on the new
runner images (recommended: run the workflow via `workflow_dispatch` with `upload=false`).

CI enforces this policy via `scripts/ci/check-gha-runner-pins.sh`, so PRs that reintroduce
`macos-latest` / `windows-latest` / `ubuntu-latest` runner labels into workflow configuration will fail.

## GitHub Actions action pinning (supply-chain hardening)

The Desktop Release workflow pins **all** third-party GitHub Actions to immutable **commit SHAs**
instead of floating tags/branches (`@v4`, `@main`, `@stable`, etc).

Why:

- Reduces the risk of compromised/malicious action updates affecting signed release artifacts.
- Makes tagged releases more reproducible (the workflow always runs the same action code).

Update guidance:

- When updating a pinned action, keep the ref as a full 40-character SHA and keep a trailing comment
  with the upstream ref (typically a **version tag** like `# v4.3.1` / `# v1`; occasionally a
  branch like `# master`) for maintainability. CI requires the comment to start with a tag/branch
  token so Dependabot can track the intended update channel.
- To resolve a tag to a commit SHA (handles annotated tags), run:

  ```bash
  git ls-remote https://github.com/<owner>/<repo> <tag> '<tag>^{}'
  ```

CI enforces this policy via `scripts/ci/check-gha-action-sha-pins.sh` (run on every PR), and the
release workflow also runs it as an early guardrail.

To keep pins fresh, `.github/dependabot.yml` enables weekly Dependabot PRs for the
`github-actions` ecosystem (grouped into a single PR).

## Testing the release pipeline (workflow_dispatch)

To test packaging/signing changes without creating a git tag, run the **Desktop Release** workflow
manually from GitHub Actions:

1. Go to **Actions → Desktop Release → Run workflow**.
2. Select the branch (or tag) you want to test.
3. Leave **upload** unchecked (default). This is **dry-run** mode:
   - bundles are built for all OS/targets
   - outputs are uploaded as **workflow artifacts** (look for artifacts starting with `desktop-bundles-...`; no GitHub Release is created/modified)
4. (Recommended) Set **version** to label the artifacts (example: `0.2.3-test`). You can also set
   **tag** (example: `v0.2.3-test`); if provided, it takes precedence.

If you set **upload=true**, the workflow will create/update a **draft** GitHub Release and attach
assets (matching the tag-driven behavior). This requires providing either `tag` or `version`.

Note: `upload=true` runs the same version validation as a tagged release (the tag/version must
match both `apps/desktop/src-tauri/tauri.conf.json` and `apps/desktop/src-tauri/Cargo.toml`), so
for ad-hoc pipeline tests without bumping the app version, prefer `upload=false`.

## Toolchain versions (keep local + CI in sync)

The release workflow pins its Node.js major via `NODE_VERSION` in `.github/workflows/release.yml`.
If you run the preflight scripts or build release bundles locally, use the same Node version to avoid
subtle differences between local and CI artifacts (see `.nvmrc` / `.node-version` / `mise.toml`).

The workflow also pins the Rust CLI tools it installs at runtime:

- `wasm-pack` (env.WASM_PACK_VERSION) – required for the `@formula/engine` WASM build
- Tauri CLI (`cargo tauri`, installed from the `tauri-cli` crate; env.TAURI_CLI_VERSION)

If you need to build release bundles locally, install the same versions (agents: use the repo cargo
wrapper):

```bash
. scripts/agent-init.sh

WASM_PACK_VERSION=0.13.1
TAURI_CLI_VERSION=2.9.5

bash scripts/cargo_agent.sh install wasm-pack --version "$WASM_PACK_VERSION" --locked --force
bash scripts/cargo_agent.sh install tauri-cli --version "$TAURI_CLI_VERSION" --locked --force
```

## Preflight validations (CI enforced)

The release workflow runs a couple of lightweight preflight scripts before it spends time building
bundles. These checks will fail the release workflow on a tagged push if the repo is not in a
releasable state.

Tip: for a one-liner that runs the key preflight checks **plus** GitHub Release asset/manifest verification for a tag, use:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z
```

Run them locally from the repo root:

```bash
# Note: CI/release workflows run these scripts under the pinned Node major
# (`NODE_VERSION` in .github/workflows/release.yml).
# Using the same major locally reduces "works locally, breaks in release" drift.
# Ensures TAURI_CLI_VERSION is patch-pinned and kept in sync with the repo's Tauri crates
# (Cargo.toml major/minor + Cargo.lock resolved patch), across workflows + docs.
node scripts/ci/check-tauri-cli-version.mjs

# Ensures the commit you're about to release has a successful run of the main CI workflow on GitHub
# (prevents tagging/releasing a commit that never passed CI).
# Requires a token with Actions read access (in GitHub Actions, this is provided via secrets.GITHUB_TOKEN).
# Accepts either GITHUB_TOKEN (preferred) or GH_TOKEN.
GITHUB_TOKEN=... node scripts/check-tag-ci-status.mjs --repo owner/repo --sha <commit>

# Ensures the tag version matches both:
# - apps/desktop/src-tauri/tauri.conf.json "version"
# - apps/desktop/src-tauri/Cargo.toml [package].version
node scripts/check-desktop-version.mjs vX.Y.Z

# Ensures packaging config includes compliance artifacts (LICENSE/NOTICE) in the
# distributed bundles (bundle.resources + Linux doc install paths).
node scripts/ci/check-desktop-compliance-artifacts.mjs

# Ensures plugins.updater.pubkey/endpoints are not placeholders and the pubkey is a valid minisign key
# when the updater is active.
node scripts/check-updater-config.mjs

# Ensures `apps/desktop/src-tauri/capabilities/*.json` only references permission identifiers
# that exist in the pinned Tauri toolchain (`cargo tauri permission ls`).
# (Requires the Tauri CLI + platform WebView dependencies; see "Toolchain versions" above.)
node scripts/check-tauri-permissions.mjs

# Ensures the packaged app registers the `formula://` URL scheme (installer/bundle metadata) and
# that `bundle.fileAssociations` includes the expected desktop file types (xlsx/xls/xlt/xla/xlsm/xltx/xltm/xlam/xlsb/csv/parquet),
# with stable MIME types (so “Open with…” / double-click integration is configured for desktop bundles).
node scripts/check-desktop-url-scheme.mjs

# Ensures the committed Cargo.lock matches the dependency graph used by the release build.
# (Fails if cargo would update Cargo.lock during the build.)
#
# Equivalent (recommended): run the same check CI uses.
bash scripts/ci/check-cargo-lock-reproducible.sh

# Under the hood, the script runs:
#
# - Workspace dependency graph (includes target-specific deps like wasm32):
#   cargo metadata --locked --format-version=1 >/dev/null
#
# - Desktop shell dependency graph (enables the `desktop` feature set):
#   cargo metadata --locked --format-version=1 --manifest-path apps/desktop/src-tauri/Cargo.toml --features desktop >/dev/null

# Ensures Windows installers will install WebView2 if it is missing.
node scripts/ci/check-webview2-install-mode.mjs

# Ensures Windows Authenticode timestamping uses HTTPS.
node scripts/ci/check-windows-timestamp-url.mjs

# Ensures Windows installers support manual rollback (downgrades) from the Releases page.
node scripts/ci/check-windows-allow-downgrades.mjs

# Ensures MSI upgrades/downgrades remain stable by pinning the WiX upgrade code.
node scripts/ci/check-windows-wix-upgrade-code.mjs

# Ensures the Tauri updater signing secrets are present for *uploading* release assets (tag pushes
# and `workflow_dispatch` runs with `upload=true`). (CI reads these from GitHub Actions secrets;
# locally requires env vars to be set.)
TAURI_PRIVATE_KEY=... node scripts/ci/check-tauri-updater-secrets.mjs
# (Optional) If you're using an encrypted PKCS#8 key instead of a minisign secret key:
TAURI_PRIVATE_KEY=... TAURI_KEY_PASSWORD=... node scripts/ci/check-tauri-updater-secrets.mjs

# (Backwards-compatible wrapper):
TAURI_PRIVATE_KEY=... node scripts/check-updater-signing-secrets.mjs

# Ensures macOS hardened-runtime entitlements include the WKWebView JIT keys
# required for JavaScript/WebAssembly to run in signed/notarized builds.
node scripts/check-macos-entitlements.mjs

# (Optional) Validate that code signing certificate secrets are valid base64 PKCS#12
# archives and decryptable with the configured password (same preflight CI runs).
APPLE_CERTIFICATE=... APPLE_CERTIFICATE_PASSWORD=... bash scripts/ci/verify-codesign-secrets.sh macos
WINDOWS_CERTIFICATE=... WINDOWS_CERTIFICATE_PASSWORD=... bash scripts/ci/verify-codesign-secrets.sh windows
```

After building (per-platform), CI also validates that the expected Tauri bundle artifacts exist
(installers/bundles, and when updater signing is enabled, the corresponding `.sig` files and
`latest.json` metadata). This catches “missing artifact” failures early (before publishing a
release):

```bash
# Run on the same OS you built on (the script defaults to `RUNNER_OS`).
# If you don't have updater signing configured locally, skip signature enforcement:
FORMULA_REQUIRE_TAURI_UPDATER_SIGNATURES=false node scripts/ci/check-desktop-release-artifacts.mjs
```

After all platform builds finish, CI also verifies the **uploaded GitHub Release assets** are
complete and consistent with the Tauri updater manifest (`latest.json`). This prevents publishing a
release where `latest.json` points at missing artifacts or missing signature files.

CI also enforces a **multi-arch safety** rule for Windows releases: when building both `x86_64` and
`aarch64` targets, the uploaded `.msi` / `.exe` installers must have **distinct filenames that
include an arch token** (for example `x64`/`x86_64`/`amd64` vs `arm64`/`aarch64`). This prevents
multi-target runs from overwriting/clobbering assets on the draft GitHub Release.

CI also enforces the same rule for Linux releases: the uploaded `.AppImage` / `.deb` / `.rpm` assets
for `x86_64` and `aarch64` must have **distinct filenames that include an arch token** (for example
`x86_64`/`amd64` vs `arm64`/`aarch64`). This prevents multi-target runs from overwriting/clobbering
assets and ensures `latest.json` can safely reference arch-specific updater payloads.

CI runs:

```bash
node scripts/verify-tauri-latest-json.mjs vX.Y.Z
node scripts/verify-tauri-updater-assets.mjs vX.Y.Z
```

You can run the same checks locally (requires a GitHub token with access to the release assets):

```bash
# Validates latest.json/latest.json.sig: required platform keys, version, signatures, and that
# platforms[*].url references real release assets.
GITHUB_REPOSITORY=owner/repo GH_TOKEN=... \
  node scripts/verify-tauri-latest-json.mjs vX.Y.Z

# Validates the full set of installer artifacts + detached signature files (<asset>.sig) on the
# GitHub Release.
GITHUB_REPOSITORY=owner/repo GH_TOKEN=... \
  node scripts/verify-tauri-updater-assets.mjs vX.Y.Z
```

If you already downloaded the manifest files (no GitHub API access needed), you can validate:

1) The required `latest.json.platforms` keys + per-platform updater artifact types (offline, **no crypto**):

```bash
node scripts/verify-tauri-latest-json.mjs --manifest latest.json --sig latest.json.sig
```

2) The updater manifest signature (offline, **cryptographic**):

```bash
node scripts/ci/verify-updater-manifest-signature.mjs latest.json latest.json.sig
```

CI also generates and uploads a `SHA256SUMS.txt` asset (SHA256 checksums for all uploaded GitHub
Release assets, including `.sig` signature files).
To reproduce locally:

```bash
# Fast path: verify `latest.json` + `latest.json.sig` + referenced assets without hashing installers/bundles.
GH_TOKEN=... node scripts/verify-desktop-release-assets.mjs --dry-run --tag vX.Y.Z --repo owner/repo

# Optional: also download and cryptographically verify updater payload signatures (slow).
GH_TOKEN=... node scripts/verify-desktop-release-assets.mjs --dry-run --verify-assets --tag vX.Y.Z --repo owner/repo

# Generate SHA256SUMS.txt for primary installers/bundles (faster; excludes `.sig` files by default).
GH_TOKEN=... node scripts/verify-desktop-release-assets.mjs --tag vX.Y.Z --repo owner/repo --out SHA256SUMS.txt

# Optional (recommended for multi-arch releases): enforce that each expected OS/arch target has at
# least one installer asset whose filename includes BOTH the version and the arch/universal token,
# and fail on any ambiguous arch-less assets that could be overwritten in multi-arch builds.
GH_TOKEN=... node scripts/verify-desktop-release-assets.mjs --tag vX.Y.Z --repo owner/repo --dry-run \
  --expectations scripts/release-asset-expectations.json

# Match CI: hash *all* release assets (including `.sig` files and updater metadata).
GH_TOKEN=... node scripts/verify-desktop-release-assets.mjs --tag vX.Y.Z --repo owner/repo --out SHA256SUMS.txt --all-assets --include-sigs

# Or, if you already downloaded the release assets into ./release-assets:
# (Note: this hashes *all* files in the directory, including `.sig` files.)
bash scripts/ci/generate-release-checksums.sh release-assets SHA256SUMS.txt
```

To inspect the required platform keys in the manifest:

```bash
jq '.platforms | keys' latest.json
```

Note: `scripts/verify-tauri-latest-json.mjs` delegates to the lower-level validator
`scripts/ci/validate-updater-manifest.mjs` when run in `<tag>` mode. It downloads `latest.json` /
`latest.json.sig` from the draft release and checks targets, signatures, and referenced assets.
The `--manifest/--sig` mode is an offline manifest structure check; use
`scripts/ci/verify-updater-manifest-signature.mjs` to cryptographically verify `latest.json.sig`.

## Compliance artifacts (LICENSE / NOTICE)

Release installers/bundles include the repo root `LICENSE` and `NOTICE` files (and any third-party
attribution content in `NOTICE`).

Where to find them after installation:

- macOS: `Formula.app/Contents/Resources/LICENSE` + `NOTICE`
- Windows: inside the installed app directory (typically `<InstallDir>\\resources\\LICENSE` +
  `<InstallDir>\\resources\\NOTICE`)
- Linux:
  - `.deb` / `.rpm`: `/usr/share/doc/formula-desktop/LICENSE` + `/usr/share/doc/formula-desktop/NOTICE`
  - `.AppImage`: inside the image at `usr/share/doc/formula-desktop/LICENSE` +
    `usr/share/doc/formula-desktop/NOTICE`

CI validates these files are present in the built bundles via:

```bash
# macOS
bash scripts/validate-macos-bundle.sh

# Linux
bash scripts/validate-linux-appimage.sh
bash scripts/validate-linux-deb.sh
bash scripts/validate-linux-rpm.sh

# Windows
pwsh -NoProfile -ExecutionPolicy Bypass -File ./scripts/validate-windows-bundles.ps1
# (Note: NSIS `.exe` payload validation uses 7-Zip / `7z.exe` when available.)
```

In addition to checking desktop integration and compliance artifacts, these validators also enforce
that the **built artifacts** report the expected **version** (and where applicable, identifier/name)
from `apps/desktop/src-tauri/tauri.conf.json`. This helps catch stale/mispackaged bundles in CI.

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
2. Update the desktop Rust crate version in `apps/desktop/src-tauri/Cargo.toml` (`[package].version`).
3. Merge the version bump to `main`.
4. Create and push a tag **with the same version** (CI enforces that the git tag matches both
   `tauri.conf.json` and `Cargo.toml`):

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
- `TAURI_KEY_PASSWORD` – optional; only used for **encrypted PKCS#8** keys.
  - Encrypted **minisign** secret keys (generated by `cargo tauri signer generate` with a password) are **not** supported by this repo’s release workflow; leave the password blank when generating minisign keys.

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
. scripts/agent-init.sh

TAURI_CLI_VERSION=2.9.5

# NOTE: Keep this version in sync with `.github/workflows/release.yml` (env.TAURI_CLI_VERSION).
bash scripts/cargo_agent.sh install tauri-cli --version "$TAURI_CLI_VERSION" --locked --force

# Generate keys (prints public + private key):
(cd apps/desktop/src-tauri && cargo tauri signer generate)
# Note: when prompted for a password, leave it blank (encrypted minisign keys are not supported by this repo's release workflow).

# Agents:
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

Code signing is **optional**. Unsigned macOS/Windows/Linux artifacts should build successfully (useful
for forks, local development, and dry-run releases).

Signing / notarization is enabled **only** when CI provides the relevant platform secrets (see below).

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
provided as a secret when available (avoids ambiguous identity selection when multiple certs exist).

Secrets used by `tauri-apps/tauri-action` (macOS build will be **unsigned** when the certificate secrets are not configured):

- `APPLE_CERTIFICATE` – base64-encoded `.p12` Developer ID certificate
- `APPLE_CERTIFICATE_PASSWORD`
- `APPLE_SIGNING_IDENTITY` – optional but recommended; example: `Developer ID Application: Your Company (TEAMID)`.
  - If set, CI will sign with this explicit identity (recommended to avoid ambiguity when multiple certs exist).
  - If missing but the certificate/password secrets are present, CI falls back to the generic `Developer ID Application` identity selector.
    (The workflow will emit a warning in this case.)
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
  - See `scripts/validate-macos-bundle.sh` (also checks basic bundle metadata like the `formula://` URL scheme and verifies the app is **universal** via `lipo`).

#### Hardened runtime entitlements (WKWebView / WASM)

The macOS app is signed with the **hardened runtime**. WKWebView (Tauri/Wry) needs explicit JIT entitlements in the signed binary so that JavaScript and WebAssembly can execute reliably.

The entitlements file used during signing is:

- `apps/desktop/src-tauri/entitlements.plist` (wired via `bundle.macOS.entitlements` in `apps/desktop/src-tauri/tauri.conf.json`)

CI guardrail (run on macOS release builds):

```bash
node scripts/check-macos-entitlements.mjs
```

This guardrail enforces:

- Required entitlements:
  - `com.apple.security.cs.allow-jit` (WKWebView/JavaScriptCore JIT)
  - `com.apple.security.cs.allow-unsigned-executable-memory` (WKWebView/JavaScriptCore executable JIT memory)
  - `com.apple.security.network.client` (outbound network; updater + HTTPS)
  - If `com.apple.security.app-sandbox` is enabled, `com.apple.security.network.server` is also required (Formula runs a loopback HTTP listener for OAuth redirects).
- Forbidden entitlements (should not be enabled for Developer ID distribution unless there is a concrete, justified need):
  - `com.apple.security.get-task-allow`
  - `com.apple.security.cs.disable-library-validation`
  - `com.apple.security.cs.disable-executable-page-protection`
  - `com.apple.security.cs.allow-dyld-environment-variables`
- Additional safety checks:
  - The entitlements plist must not contain any unexpected keys (keep the signed entitlement surface minimal; allowlisted keys are the ones above, plus `com.apple.security.app-sandbox` when enabled).
  - All entitlements must be boolean `<true/>` values (no `<false/>` / non-boolean payloads).

Release workflow note: when macOS signing secrets are configured, CI extracts the entitlements from the built `.app` (`codesign -d --entitlements :-`) and validates them with `node scripts/check-macos-entitlements.mjs`. This ensures the entitlements are actually embedded in the signed bundle (protects against config drift where the plist exists but isn’t used during signing).

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
   app="$(find apps/desktop/src-tauri/target -maxdepth 8 -type d -path '*/release/bundle/macos/*.app' 2>/dev/null | head -n 1 || true)"
   test -n "$app" || { echo "No .app bundle found under apps/desktop/src-tauri/target/release/bundle/macos/*.app (or apps/desktop/src-tauri/target/*/release/bundle/macos/*.app)" >&2; exit 1; }
   echo "Checking app at: $app"
   codesign --verify --deep --strict --verbose=2 "$app"
   codesign -d --entitlements :- "$app" 2>&1 | grep -E "allow-jit|allow-unsigned-executable-memory|network\\.client"
   spctl --assess --type execute -vv "$app"

   # (Optional) If you're testing a notarized/stapled CI artifact:
   xcrun stapler validate -v "$app"
   ```

4. Launch the app and sanity-check runtime behavior:
    - The window should render (no blank WebView).
    - Network features work (e.g. updater check / HTTPS fetches).
    - Cross-origin isolation still works in the packaged app (see `pnpm -C apps/desktop check:coi`).

For CI-style bundle verification (DMG mount + Info.plist sanity + universal `lipo` check + optional codesign/notarization checks), you can also run:

```bash
bash scripts/validate-macos-bundle.sh
```

Note: this repo’s **release** artifacts are built as **universal** macOS bundles. If you’re building
locally and want this script to pass, build with the universal target:

```bash
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build --target universal-apple-darwin
```

#### Troubleshooting: blank window / crashes in a signed build

If a signed/notarized build launches with a blank window or crashes immediately, check:

1. The **entitlements actually embedded in the signed app** (not just the plist file in the repo):

   ```bash
   codesign -d --entitlements :- /Applications/Formula.app
   ```

   Ensure it includes:
   - `com.apple.security.cs.allow-jit`
   - `com.apple.security.cs.allow-unsigned-executable-memory`
   - `com.apple.security.network.client` (outbound network; enforced when sandboxing is enabled)

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
  # If you built without --target, bundles may land under src-tauri\\target\\release\\bundle\\...
  signtool verify /pa /all /v apps\desktop\src-tauri\target\release\bundle\nsis\*.exe
  signtool verify /pa /all /v apps\desktop\src-tauri\target\release\bundle\msi\*.msi
 
  # If you built with --target (as CI does), bundles land under src-tauri\\target\\<triple>\\release\\bundle\\...
  signtool verify /pa /all /v apps\desktop\src-tauri\target\x86_64-pc-windows-msvc\release\bundle\nsis\*.exe
  signtool verify /pa /all /v apps\desktop\src-tauri\target\x86_64-pc-windows-msvc\release\bundle\msi\*.msi
  signtool verify /pa /all /v apps\desktop\src-tauri\target\aarch64-pc-windows-msvc\release\bundle\nsis\*.exe
  signtool verify /pa /all /v apps\desktop\src-tauri\target\aarch64-pc-windows-msvc\release\bundle\msi\*.msi
  ```
  - Release CI runs `pwsh -NoProfile -ExecutionPolicy Bypass -File scripts/validate-windows-bundles.ps1` on the produced installers when
    `WINDOWS_CERTIFICATE` is configured. This validates that the expected installers were produced
    and that they are Authenticode-signed and timestamped.

Timestamping:

- Authenticode timestamping embeds a trusted signing time in the signature so it remains valid after
  the code signing certificate expires.
- The Authenticode timestamp server is configured in `apps/desktop/src-tauri/tauri.conf.json` under
  `bundle.windows.timestampUrl` (currently `https://timestamp.digicert.com`).
  - Release CI preflight enforces this uses HTTPS to avoid sending/receiving timestamp requests over
    insecure HTTP and to reduce the risk of MITM tampering (see
    `scripts/ci/check-windows-timestamp-url.mjs`).
- If a release fails due to timestamping/network issues, switch `timestampUrl` to another **HTTPS**
  timestamp server provided/recommended by your signing certificate vendor and re-run the workflow.
  - For a one-off fallback without committing a config change, re-run the release workflow via
    **Actions → Desktop Release → Run workflow** and set the `windows_timestamp_url` input (must be
    `https://...`).

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

Note: some environments (including GitHub-hosted runners) can have **multiple** Windows SDK
versions installed, and the newest SDK version may be missing the ARM64 lib subset. CI detects a
SDK version that actually contains both `um\arm64` and `ucrt\arm64` and pins the MSVC environment
setup to that SDK version so cross-compiles link reliably.

When cross-compiling locally from an x64 Windows machine, run the build in a Visual Studio
Developer Prompt configured for **amd64 → arm64** (CI uses `ilammy/msvc-dev-cmd` with `arch:
amd64_arm64`).

For example, in a Developer Command Prompt you can run:

```powershell
vcvarsall.bat amd64_arm64
```

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

Expected output includes `AA64` (for example: `AA64 machine (ARM64)`).

Sanity-check (optional): verify the produced MSI targets ARM64 (Template Summary contains `ARM64`):

```powershell
$installer = New-Object -ComObject WindowsInstaller.Installer
Get-ChildItem apps\desktop\src-tauri\target\aarch64-pc-windows-msvc\release\bundle\msi\*.msi | ForEach-Object {
  $db = $installer.OpenDatabase($_.FullName, 0)
  $summary = $db.SummaryInformation(0)
  $template = $summary.Property(7)
  Write-Host "$($_.Name): template=$template"
  if ($template -notmatch '(?i)arm64') {
    throw "Expected MSI Template Summary to include ARM64 for $($_.FullName), got: $template"
  }
}
```

GitHub-hosted runner images do not always include this workload by default. The release workflow
checks for a complete ARM64 MSVC + SDK toolchain:

- MSVC: `VC\\Tools\\MSVC\\*\\lib\\arm64` + `VC\\Tools\\MSVC\\*\\bin\\Hostx64\\arm64\\{cl.exe,link.exe}`
- Windows SDK: `Windows Kits\\10\\Lib\\*\\{um,ucrt}\\arm64`

If any of these are missing, CI installs the MSVC ARM64 component via `vs_installer.exe` and fails
with a clear error if the runner image still lacks the required ARM64 SDK libraries. If the Windows
SDK ARM64 libs are missing, CI also attempts to install the matching `Windows10SDK.*` component via
`vs_installer.exe` before failing (tries both `Windows10SDK.*` and `Windows11SDK.*` component IDs).

CI smoke test:

- `.github/workflows/windows-arm64-smoke.yml` runs `cargo tauri build --target aarch64-pc-windows-msvc`
  and asserts that the expected Windows bundles land under
  `apps/desktop/src-tauri/target/aarch64-pc-windows-msvc/release/bundle/**`.
  - It also runs on a weekly schedule in the upstream repo to catch `windows-2022` runner image
    regressions (missing MSVC/SDK ARM64 components) early.
  - Additional validations:
    - `formula-desktop.exe` must report **AA64** in `dumpbin /headers` output (guards against accidentally producing an x64 binary).
    - The ARM64 MSI installer must have `ARM64` in its **Template Summary** (guards against producing a non-ARM64 MSI).
    - Windows installers must bundle/reference the WebView2 runtime bootstrapper/runtime (see `scripts/ci/check-windows-webview2-installer.py`).

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
    - Note: for `downloadBootstrapper` mode, installers may reference the bootstrapper via a download URL (e.g.
      `go.microsoft.com/fwlink/?LinkId=2124703`) instead of embedding `MicrosoftEdgeWebView2Setup.exe`; the checker treats
      those URL markers as valid evidence too.

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
by `tauri-apps/tauri-action` (pinned to an immutable commit SHA in `.github/workflows/release.yml`;
currently v0.6.1):

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

#### Templated endpoints (`{{target}}`, `{{current_version}}`)

Tauri's updater supports **templated endpoint URLs**. At runtime it replaces:

- `{{target}}` — the current Tauri updater **target string** (used to select the right OS/arch artifact)
- `{{current_version}}` — the currently-installed app version

This is useful when hosting update metadata outside GitHub (or when you want per-target/per-version
paths on a CDN).

Example **placeholder** endpoint (do not ship this; it is intentionally caught by CI as a placeholder):

```
https://releases.formula.app/{{target}}/{{current_version}}
```

Replace it with your real update JSON URL(s) before tagging a release. CI enforces this via
`scripts/check-updater-config.mjs` when `plugins.updater.active=true`.

Note: `scripts/check-updater-config.mjs` also treats obvious placeholders like `example.com`,
`localhost`, and URLs containing `REPLACE_WITH...` as invalid on tagged releases.

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
    "darwin-x86_64": { "url": "…", "signature": "…" },
    "darwin-aarch64": { "url": "…", "signature": "…" },
    "windows-x86_64": { "url": "…", "signature": "…" },
    "windows-aarch64": { "url": "…", "signature": "…" },
    "linux-x86_64": { "url": "…", "signature": "…" },
    "linux-aarch64": { "url": "…", "signature": "…" }
  }
}
```

Expected `{{target}}` / `latest.json.platforms` keys for this repo’s **tagged release** matrix (CI
enforced; see `docs/desktop-updater-target-mapping.md`):

- **macOS (universal build):** `darwin-x86_64` and `darwin-aarch64` → macOS updater tarball
  (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`).
- **Windows x64:** `windows-x86_64` → updater installer (currently the **`.msi`**).
- **Windows ARM64:** `windows-aarch64` → updater installer (currently the **`.msi`**).
- **Linux x86_64:** `linux-x86_64` → updater payload (typically the `.AppImage`).
- **Linux ARM64:** `linux-aarch64` → updater payload (typically the `.AppImage`).

Note: `apps/desktop/src-tauri/tauri.conf.json` sets `bundle.targets: "all"`, which enables all
supported bundlers for the current platform (including **MSI/WiX** + **NSIS** on Windows). CI still
passes `--bundles msi,nsis` and installs WiX + NSIS explicitly so Windows releases always include
both installer formats.

For reference, this is how the release workflow’s Tauri build targets map to updater targets:

| Workflow build | Tauri build args | Rust target triple | `latest.json` platform key(s) |
| --- | --- | --- | --- |
| macOS universal | `--target universal-apple-darwin` | `aarch64-apple-darwin` + `x86_64-apple-darwin` | `darwin-x86_64` + `darwin-aarch64` |
| Windows x64 | `--target x86_64-pc-windows-msvc --bundles msi,nsis` | `x86_64-pc-windows-msvc` | `windows-x86_64` |
| Windows ARM64 | `--target aarch64-pc-windows-msvc --bundles msi,nsis` | `aarch64-pc-windows-msvc` | `windows-aarch64` |
| Linux x86_64 | `--bundles appimage,deb,rpm` | `x86_64-unknown-linux-gnu` | `linux-x86_64` |
| Linux ARM64 | `--bundles appimage,deb,rpm` | `aarch64-unknown-linux-gnu` | `linux-aarch64` |

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
- openSUSE: `zypper info libwebkit2gtk-4_1-0`
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

- `shared-mime-info` – provides `update-mime-database`, required for installing our packaged
  shared-mime-info definitions (notably Parquet) and making file associations resolve reliably.
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

Recommended (repo script; runs static `dpkg-deb` metadata checks + optional installability check in an Ubuntu container):

```bash
# Auto-discovers DEB(s) under <target>/release/bundle/deb/*.deb (or <target>/<triple>/release/bundle/deb/*.deb)
bash scripts/validate-linux-deb.sh

# Or validate a specific .deb file:
deb_pkg="$(find apps/desktop/src-tauri/target apps/desktop/target target -type f -path '*/release/bundle/deb/*.deb' 2>/dev/null | head -n 1 || true)"
test -n "$deb_pkg" || { echo "No .deb found under any target root (expected <target>/release/bundle/deb/*.deb or <target>/<triple>/release/bundle/deb/*.deb)" >&2; exit 1; }
bash scripts/validate-linux-deb.sh --deb "$deb_pkg"

# Skip the container step (static checks only):
bash scripts/validate-linux-deb.sh --no-container
```

From `apps/desktop/src-tauri`:

```bash
# Inspect the control file (check Depends: ...)
deb="$(find apps/desktop/src-tauri/target apps/desktop/target target -type f -path '*/release/bundle/deb/*.deb' 2>/dev/null | head -n 1 || true)"
test -n "$deb" || { echo "No .deb found under any target root (expected <target>/release/bundle/deb/*.deb or <target>/<triple>/release/bundle/deb/*.deb)" >&2; exit 1; }
dpkg -I "$deb"

# Extract and confirm all linked shared libraries resolve
tmpdir="$(mktemp -d)"
dpkg-deb -x "$deb" "$tmpdir"
ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
```

For a clean install test (no GUI required), use a container:

```bash
deb_dir="$(cd "$(dirname "$deb")" && pwd -P)"
docker run --rm -it \
  -v "$deb_dir:/deb:ro" \
  ubuntu:24.04 bash -lc '
    apt-get update
    apt-get install -y --no-install-recommends /deb/*.deb
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

Note: the container install smoke test runs on the **same CPU architecture as the host**. Validate
ARM64 `.deb` packages on an ARM64 host/runner (or use QEMU/binfmt emulation).

To run the same install smoke test CI uses (includes Docker platform pinning and extra diagnostics), run:

```bash
bash scripts/ci/linux-package-install-smoke.sh deb
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

- `shared-mime-info` – provides `update-mime-database`, required for installing our packaged
  shared-mime-info definitions (notably Parquet) and making file associations resolve reliably.
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

Recommended (repo script; runs static RPM queries + an installability check in a Fedora container):

```bash
# Auto-discovers RPM(s) under <target>/release/bundle/rpm/*.rpm (or <target>/<triple>/release/bundle/rpm/*.rpm)
bash scripts/validate-linux-rpm.sh

# Or validate a specific .rpm file:
rpm_pkg="$(find apps/desktop/src-tauri/target apps/desktop/target target -type f -path '*/release/bundle/rpm/*.rpm' 2>/dev/null | head -n 1 || true)"
test -n "$rpm_pkg" || { echo "No RPM found under any target root (expected <target>/release/bundle/rpm/*.rpm or <target>/<triple>/release/bundle/rpm/*.rpm)" >&2; exit 1; }
bash scripts/validate-linux-rpm.sh --rpm "$rpm_pkg"

# Skip the Fedora container step (static checks only):
bash scripts/validate-linux-rpm.sh --no-container
```

Note: `--no-container` still extracts the RPM payload to validate `.desktop` file `MimeType=` entries
(file association metadata), so it requires `rpm2cpio` + `cpio` on the host.

Note: the Fedora container install smoke test runs on the **same CPU architecture as the host**
(x86_64 Fedora image on x86_64 hosts; aarch64 Fedora image on aarch64 hosts). Validate ARM64 RPMs on
an ARM64 host/runner (or use QEMU/binfmt emulation), or use `--no-container` for metadata-only
checks.

From `apps/desktop/src-tauri`:

```bash
# Inspect declared dependencies (check webkit2gtk/gtk3/appindicator/etc)
rpm_pkg="$(find apps/desktop/src-tauri/target apps/desktop/target target -type f -path '*/release/bundle/rpm/*.rpm' 2>/dev/null | head -n 1 || true)"
test -n "$rpm_pkg" || { echo "No RPM found under any target root (expected <target>/release/bundle/rpm/*.rpm or <target>/<triple>/release/bundle/rpm/*.rpm)" >&2; exit 1; }
rpm -qpR "$rpm_pkg"

# Extract and confirm all linked shared libraries resolve
# (requires `cpio`: Fedora `dnf -y install cpio`, Debian/Ubuntu `apt-get install -y cpio`)
tmpdir="$(mktemp -d)"
# `rpm_pkg` is usually a relative path; run `rpm2cpio` from the current directory and only
# `cd` for the extraction destination.
rpm2cpio "$rpm_pkg" | (cd "$tmpdir" && cpio -idmv)
ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
```

For a clean install test (no GUI required), use a Fedora container:

```bash
rpm_dir="$(cd "$(dirname "$rpm_pkg")" && pwd -P)"
docker run --rm -it \
  -v "$rpm_dir:/rpm:ro" \
  fedora:40 bash -lc '
    # The Tauri updater `.sig` files are *not* RPM GPG signatures, so install with --nogpgcheck.
    dnf -y install --nogpgcheck --setopt=install_weak_deps=False /rpm/*.rpm
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

To run the same install smoke test CI uses (Fedora by default; includes Docker platform pinning), run:

```bash
bash scripts/ci/linux-package-install-smoke.sh rpm
```

Optional: openSUSE smoke install (helps validate our RPM rich-deps cover openSUSE package naming):

```bash
# Use the same CI helper script, but override the container image.
FORMULA_RPM_SMOKE_IMAGE=opensuse/tumbleweed:latest bash scripts/ci/linux-package-install-smoke.sh rpm
```

CI guardrails (tagged releases):

- `bash scripts/ci/verify-linux-package-deps.sh` inspects the produced `.rpm` with `rpm -qpR` and fails the workflow if the
  **core runtime dependencies** are missing from the RPM metadata.
- `bash scripts/ci/linux-package-install-smoke.sh rpm` installs the `.rpm` into a clean Fedora container and fails if
  `ldd /usr/bin/formula-desktop` reports missing shared libraries.
- Optional: set the GitHub Actions variable `FORMULA_OPENSUSE_SMOKE=1` to also run an openSUSE/Tumbleweed RPM install smoke
  test (helps validate the RPM rich-deps cover openSUSE package naming).

Note: showing a tray icon also requires a desktop environment with **StatusNotifier/AppIndicator**
support (e.g. the GNOME Shell “AppIndicator and KStatusNotifierItem Support” extension).

## Linux: `.AppImage` validation

The AppImage is a **self-contained SquashFS bundle**. Before publishing a release, validate that it
contains the expected payload (binary + resources) and that the embedded `.desktop` file declares
the expected MIME/file associations.

Recommended (repo scripts):

```bash
# Auto-discovers AppImage(s) under <target>/release/bundle/appimage/*.AppImage (or <target>/<triple>/release/bundle/appimage/*.AppImage)
bash scripts/validate-linux-appimage.sh

# Optional: also run the CI smoke test locally (checks ELF arch + ldd "not found", no GUI).
bash scripts/ci/check-appimage.sh
```

CI note: the release workflow also runs a lightweight smoke test that validates AppImage extraction
+ELF architecture + `ldd` (no GUI): `bash scripts/ci/check-appimage.sh`.

Note: AppImage validation executes the AppImage to extract it (`--appimage-extract`), so the
AppImage’s CPU architecture must match the host (x86_64 vs aarch64) unless you are using emulation.

Note: extraction (`--appimage-extract`) typically relies on `unsquashfs` (package: `squashfs-tools`).
If AppImage extraction fails with an `unsquashfs`/SquashFS error, install `squashfs-tools` and retry.

Manual inspection (useful when debugging bundling issues):

```bash
appimage="$(find apps/desktop/src-tauri/target apps/desktop/target target -type f -path '*/release/bundle/appimage/*.AppImage' 2>/dev/null | head -n 1 || true)"
test -n "$appimage" || { echo "No AppImage found under any target root (expected <target>/release/bundle/appimage/*.AppImage or <target>/<triple>/release/bundle/appimage/*.AppImage)" >&2; exit 1; }
appimage="$(realpath "$appimage")"
chmod +x "$appimage"

tmpdir="$(mktemp -d)"
(cd "$tmpdir" && "$appimage" --appimage-extract >/dev/null)
root="$tmpdir/squashfs-root"

# Confirm the main payload exists
test -x "$root/usr/bin/formula-desktop"

# Confirm a desktop entry exists and includes MIME types (file associations)
desktop_file="$(ls "$root/usr/share/applications/"*.desktop 2>/dev/null | head -n 1 || true)"
test -n "$desktop_file" || { echo "No .desktop file found under $root/usr/share/applications/" >&2; exit 1; }
cat "$desktop_file"
grep -E '^MimeType=' "$desktop_file"

# Optional: validate .desktop syntax (requires `desktop-file-utils`)
command -v desktop-file-validate >/dev/null && desktop-file-validate "$desktop_file" || true

# Optional (recommended): run the same desktop-integration validator CI uses for Linux packages.
# This checks `MimeType=` coverage and that `Exec=` includes a %u/%U/%f/%F placeholder.
python3 scripts/ci/verify_linux_desktop_integration.py --package-root "$root"
```

## 5) Verifying a release

After the workflow completes, open the GitHub Release (draft) and confirm the expected artifacts
are attached:

Note: the in-app updater downloads whatever URLs `latest.json` points at (per-platform). The
auto-update artifact is not always the same file you’d choose for manual install (see
`docs/desktop-updater-target-mapping.md`):

- macOS: updater uses a tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) (not the `.dmg`)
- Linux: updater uses `*.AppImage` (not `.deb`/`.rpm`)
- Windows: updater uses the **`.msi`** installer referenced in `latest.json` (the NSIS `.exe` is shipped for manual install/downgrade)

Quick reference (auto-update vs manual install):

| Target key (`latest.json.platforms`) | Auto-update asset (`platforms[key].url`) | Manual install |
| --- | --- | --- |
| `darwin-x86_64` | updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) | `.dmg` |
| `darwin-aarch64` | updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) | `.dmg` |
| `windows-x86_64` | `*.msi` (Windows Installer; updater runs this) | `.msi` / `.exe` (NSIS) |
| `windows-aarch64` | `*.msi` (Windows Installer; updater runs this) | `.msi` / `.exe` (NSIS) |
| `linux-x86_64` | `*.AppImage` | `.deb` / `.rpm` (AppImage optional) |
| `linux-aarch64` | `*.AppImage` | `.deb` / `.rpm` (AppImage optional) |

### Verify updater assets (auto-update) — quick checklist

Before validating in-app auto-update behavior, confirm the updater manifest (`latest.json`) is
wired to the correct **updater-consumed** artifacts:

Note: for end-to-end **in-app updater** testing, prefer installing a prior version using the same
format the updater will apply:

- **Windows:** install via the **WiX `.msi`** (the updater runs the MSI referenced in `latest.json`).
  The NSIS `.exe` is shipped for manual install/downgrade and may result in side-by-side installs if
  you later apply an MSI update.
- **Linux:** install/run the **`.AppImage`** (the updater payload). `.deb`/`.rpm` are intended for
  manual install/updates via your package manager.
- **macOS:** install via the `.dmg` (copy to `/Applications`); the updater downloads the tarball
  payload and replaces the app bundle.

1. Confirm the Release contains:
   - `latest.json`
   - `latest.json.sig` (signature for `latest.json`)
2. Download `latest.json` and inspect the platform URLs:
   - Published “latest” (public):
      - `curl -L -o latest.json https://github.com/OWNER/REPO/releases/latest/download/latest.json`
      - `curl -L -o latest.json.sig https://github.com/OWNER/REPO/releases/latest/download/latest.json.sig`
   - Specific tag (useful for draft QA; may require auth if the release is still a draft):
      - Recommended (uses the GitHub API via `gh`; works for draft releases when authenticated):
        - `GH_TOKEN=... gh release download vX.Y.Z --repo OWNER/REPO --pattern 'latest.json*'`
      - `curl -L -o latest.json https://github.com/OWNER/REPO/releases/download/vX.Y.Z/latest.json`
      - `curl -L -o latest.json.sig https://github.com/OWNER/REPO/releases/download/vX.Y.Z/latest.json.sig`
   - Inspect:
      - `jq '.platforms | keys' latest.json`
      - `jq -r '.platforms | to_entries[] | "\(.key)\t\(.value.url)"' latest.json`
      - Confirm each platform entry has a non-empty `signature` string (per-payload updater signature; distinct from `latest.json.sig`):
        - `jq -r '.platforms | to_entries[] | select((.value.signature // "") == "") | .key' latest.json`
3. Confirm each `platforms[*].url` points at the expected **updater** asset type (not a manual-only installer):
   - macOS: updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) (**not** `.dmg`)
   - Windows: `*.msi` (CI expects the manifest to reference the MSI; the `.exe` is for manual install)
   - Linux: `*.AppImage` (**not** `.deb`/`.rpm`)
   - Multi-arch correctness:
     - macOS: it is normal for `darwin-x86_64` and `darwin-aarch64` to point at the **same** universal updater tarball.
     - Windows: `windows-x86_64` and `windows-aarch64` should point at **different** `.msi` files whose filenames include an arch token (e.g. `x64`/`x86_64`/`amd64` vs `arm64`/`aarch64`).
     - Linux: `linux-x86_64` and `linux-aarch64` should point at **different** `.AppImage` files whose filenames include an arch token (e.g. `x86_64`/`amd64` vs `arm64`/`aarch64`).
4. Confirm each URL filename matches an actual Release asset (no broken/missing assets).
   - This repo also expects detached signature assets (`<asset>.sig`) to be uploaded alongside each updater payload.
5. (Optional) Verify the manifest signature locally:
   - `node scripts/ci/verify-updater-manifest-signature.mjs latest.json latest.json.sig`
   - Note: this uses the updater public key embedded in `apps/desktop/src-tauri/tauri.conf.json`.
6. (Optional) Verify updater payload signatures (slow; downloads the updater assets referenced in `latest.json`):
   - `GITHUB_TOKEN=... node scripts/verify-desktop-release-assets.mjs --tag vX.Y.Z --repo OWNER/REPO --dry-run --verify-assets`
7. (Optional) Verify downloaded updater assets against `SHA256SUMS.txt`.

### One-liner: release smoke test

To run the repo’s release sanity checks (version check, updater config validation, compliance artifact config validation, and GitHub Release asset/manifest verification) in one command:

```bash
# Requires a GitHub token (or pass --token).
#
# Note: --repo defaults to $GITHUB_REPOSITORY (in CI) or your git remote "origin" (if it's a GitHub URL),
# so you can usually omit it for the upstream repo.
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z
```

Notes:

- `GH_TOKEN` is also supported (same value as `GITHUB_TOKEN`).
- For GitHub Enterprise, set `GITHUB_API_URL=https://your-ghe.example/api/v3` and pass `--repo owner/name` if needed.

If you want a faster check that **does not** download/hash all release bundles (still validates `latest.json` + signatures + asset presence), add `--dry-run`:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z --dry-run
```

To additionally verify the **updater payload signatures** (slow; downloads the updater bundles referenced in `latest.json`), add `--verify-assets`:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z --dry-run --verify-assets
```

To additionally enforce that the release contains **all expected OS/arch targets** (recommended for multi-arch releases), pass the expectations config:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z --dry-run \
  --expectations scripts/release-asset-expectations.json
```

To also check that the draft release includes the supply-chain metadata assets (SBOM + provenance bundles), add:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z --dry-run --check-supply-chain
```

To **fail** the smoke test if supply-chain metadata assets are missing, use:

```bash
GITHUB_TOKEN=... node scripts/release-smoke-test.mjs --tag vX.Y.Z --dry-run --require-supply-chain
```

If you have locally-built Tauri bundles and want to run any platform-specific bundle validators too:

```bash
node scripts/release-smoke-test.mjs --tag vX.Y.Z --local-bundles
```

1. Open the GitHub Release (draft) and confirm:
     - Updater metadata: `latest.json` and `latest.json.sig`
     - `SHA256SUMS.txt` (SHA256 checksums for all release assets)
     - SBOM: `sbom.spdx.json` (SPDX JSON; Rust + JS dependency set; also uploaded as a workflow artifact named `sbom-<tag>`)
     - Build provenance bundles: `provenance-*.intoto.jsonl` (also uploaded as workflow artifacts `provenance-*`)
       - Expected (one per target): `provenance-universal-apple-darwin.intoto.jsonl`, `provenance-x86_64-pc-windows-msvc.intoto.jsonl`, `provenance-aarch64-pc-windows-msvc.intoto.jsonl`, `provenance-x86_64-unknown-linux-gnu.intoto.jsonl`, `provenance-aarch64-unknown-linux-gnu.intoto.jsonl`
     - Note: GitHub also shows auto-generated “Source code (zip)” / “Source code (tar.gz)” entries; those are **not** installers or updater payloads.
     - macOS (**universal**): `.dmg` (installer) + updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`)
     - Windows **x64**: installers (WiX `.msi` **and** NSIS `.exe`, filename typically includes `x64` / `x86_64`)
     - Windows **ARM64**: installers (WiX `.msi` **and** NSIS `.exe`, filename typically includes `arm64` / `aarch64`)
     - Linux (**x86_64 + ARM64**): `.AppImage` + `.deb` + `.rpm` for each architecture (filenames typically include `x86_64` / `amd64` vs `arm64` / `aarch64`)

   This repo requires Tauri updater signing for tagged releases, so expect `.sig` signature files to
   be uploaded alongside the produced artifacts:
   - macOS: `.dmg.sig` and updater tarball signature (`*.tar.gz.sig`/`*.tgz.sig`; usually `*.app.tar.gz.sig`)
   - Windows (each architecture): `.msi.sig` and `.exe.sig`
   - Linux (each architecture): `.AppImage.sig`, `.deb.sig`, `.rpm.sig`

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

    Note: SBOM/provenance publishing is warnings-only initially. If `sbom.spdx.json` or
    `provenance-*.intoto.jsonl` are missing from the draft release, check the workflow artifacts for
    `sbom-*` / `provenance-*` bundles and the workflow logs for upload warnings.

    Optional: to enforce SBOM/provenance presence on tagged releases, set the GitHub Actions variable
    `FORMULA_ENFORCE_SUPPLY_CHAIN_ASSETS=1`.

    Build provenance: the release workflow also generates **build provenance attestations** for the
    uploaded desktop artifacts. You can view them in GitHub’s **Attestations** UI for the workflow
    run, and the raw attestation bundles are attached to the draft GitHub Release as
    `provenance-*.intoto.jsonl` (and also uploaded as workflow artifacts `provenance-*`).

    Quick verify (CLI): after downloading an installer/bundle from the release, you can verify its
    GitHub attestation with the GitHub CLI:

    ```bash
    # Example (replace with a specific downloaded asset file):
    gh attestation verify path/to/Formula-Setup-x64.msi --repo wilson-anysphere/formula
    ```

    To verify using the **provenance bundle attached to the release** (without fetching attestations
    from the GitHub API), download the matching `provenance-*.intoto.jsonl` asset and pass it via
    `--bundle`:

    ```bash
    gh attestation verify path/to/Formula-Setup-x64.msi --repo wilson-anysphere/formula \
      --bundle provenance-x86_64-pc-windows-msvc.intoto.jsonl
    ```

    Note: `gh attestation verify` still verifies against Sigstore trust roots/certificates; the
    `provenance-*.intoto.jsonl` files are the raw Sigstore bundles produced during the workflow run.
2. Download `latest.json` and confirm `platforms` includes entries for:
   - `darwin-x86_64` (macOS Intel; points at the updater tarball)
   - `darwin-aarch64` (macOS Apple Silicon; points at the updater tarball)
   - `windows-x86_64` (Windows x64)
   - `windows-aarch64` (Windows ARM64)
   - `linux-x86_64` (Linux x86_64)
   - `linux-aarch64` (Linux ARM64)

   Note: tagged-release CI is intentionally **strict** about the required `{os}-{arch}` platform key
   names (see `docs/desktop-updater-target-mapping.md`). `latest.json` may also include additional
   installer-specific keys, and CI validates those entries reference real release assets. If a
   Tauri/tauri-action upgrade changes the required `{os}-{arch}` key set, CI is expected to fail
   loudly with an expected vs actual diff so we update the docs + validators together.

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
   - `darwin-*` → updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`) (**not** `.dmg`)
   - `windows-*` → `*.msi` (updater runs the Windows Installer; this repo requires the manifest to reference the MSI)
   - `linux-*` → `*.AppImage`

3. Download the artifacts and do quick sanity checks:

   ### macOS: confirm the app is universal

    Run `lipo -info` on the bundled executable (`Formula.app/Contents/MacOS/formula-desktop`):

    ```bash
    # Option A: from the updater tarball (`*.app.tar.gz` preferred; allow `*.tar.gz`/`*.tgz`).
    #
    # Prefer selecting the exact updater tarball referenced in latest.json so we don't accidentally
    # pick:
    # - GitHub "Source code (tar.gz)" archives
    # - Linux `.AppImage.tar.gz` bundles
    app_tgz=""
    if [[ -f latest.json ]] && command -v jq >/dev/null 2>&1; then
      app_tgz="$(
        jq -r '.platforms["darwin-aarch64"].url // .platforms["darwin-x86_64"].url // empty' latest.json \
          | sed 's|[?#].*$||' \
          | sed 's|.*/||' \
          | head -n 1
      )"
      if [[ -n "$app_tgz" && ! -f "$app_tgz" ]]; then
        app_tgz=""
      fi
    fi
    if [[ -z "$app_tgz" ]]; then
      shopt -s nullglob
      for candidate in *.app.tar.gz *.app.tgz *.tar.gz *.tgz; do
        # Avoid confusing macOS tarballs with Linux `.AppImage.tar.gz` bundles (sometimes downloaded
        # into the same folder during release QA).
        if echo "$candidate" | grep -q -i -F '.appimage.'; then
          continue
        fi
        app_tgz="$candidate"
        break
      done
      shopt -u nullglob
    fi
    test -n "$app_tgz" || { echo "No macOS updater tarball found (see latest.json or expected *.app.tar.gz/*.app.tgz/*.tar.gz/*.tgz)" >&2; exit 1; }
    tar -xzf "$app_tgz"
    lipo -info "Formula.app/Contents/MacOS/formula-desktop"

    # Expected output includes both: x86_64 arm64
    ```

    If you only have a `.dmg`, mount it and inspect the `.app` inside:

    ```bash
    dmg="$(ls *.dmg 2>/dev/null | head -n 1 || true)"
    test -n "$dmg" || { echo "No .dmg file found in the current directory" >&2; exit 1; }
    mnt="$(mktemp -d)"
    hdiutil attach "$dmg" -nobrowse -mountpoint "$mnt"
    lipo -info "$mnt/Formula.app/Contents/MacOS/formula-desktop"
    hdiutil detach "$mnt"
    ```

    Tip: on macOS, you can run the repo helper to validate the DMG (and the updater tarball if
    present), including a universal `lipo` check:

    ```bash
    bash scripts/validate-macos-bundle.sh --dmg "$dmg"
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
   signtool verify /pa /all /v path\to\installer.exe
   signtool verify /pa /all /v path\to\installer.msi
   ```

   ### Windows: WebView2 install smoke test (clean VM)

   On a clean Windows VM **without** WebView2 (or after uninstalling **Microsoft Edge WebView2 Runtime**),
   run the installer. It should install WebView2 via the configured Evergreen bootstrapper and then
   the app should launch successfully. (This requires an internet connection when using the bootstrapper modes.)

   ### Linux: inspect dependencies + `ldd` smoke check

   ```bash
   # Dependency metadata (ensure the runtime deps are present)
   # (Pick the package files matching the architecture you're validating: amd64/x86_64 vs arm64/aarch64.)
   deb="$(ls *.deb 2>/dev/null | head -n 1 || true)"
   rpm="$(ls *.rpm 2>/dev/null | head -n 1 || true)"
   test -n "$deb" || { echo "No .deb file found in the current directory" >&2; exit 1; }
   test -n "$rpm" || { echo "No .rpm file found in the current directory" >&2; exit 1; }

   dpkg -I "$deb"
   rpm -qpi "$rpm"
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
   shasum -a 256 *.dmg

   # Linux
   sha256sum *.AppImage
   ```

   ```powershell
   # Windows (PowerShell)
   Get-FileHash -Algorithm SHA256 .\*.msi
   ```

Also verify **cross-origin isolation** is enabled in the packaged app (required for `SharedArrayBuffer` and the Pyodide Worker backend):

- From source (recommended preflight): `pnpm -C apps/desktop check:coi`
  - If you already ran a Tauri build (for example `cargo tauri build` / `tauri-action`), you can reuse the built artifacts:
    `pnpm -C apps/desktop check:coi -- --no-build`
  - Linux/CI note: if the check hangs in a headless environment, set `FORMULA_COI_TIMEOUT_SECS=<seconds>` to apply an outer timeout
    (set it to `0` to disable).
- Or in an installed build: ensure there is no startup toast complaining about missing cross-origin isolation, and (if you have DevTools)
   confirm `globalThis.crossOriginIsolated === true`.

### File associations + deep link scheme (CI guardrailed)

The release workflow also inspects the **built artifacts** to ensure OS
integration metadata made it into the final bundles (not just `tauri.conf.json`):

- macOS: the built app bundle `Info.plist` registers the configured file associations
  (`bundle.fileAssociations` from `tauri.conf.json`) via `CFBundleDocumentTypes`, and registers the
  `formula` deep-link scheme via `CFBundleURLTypes`.
- Windows: the built installers register:
  - file associations from `apps/desktop/src-tauri/tauri.conf.json` (`bundle.fileAssociations`)
  - the `formula://` URL protocol handler
  - Note: MSI validation is static/authoritative (Windows Installer tables via COM). NSIS `.exe`
    validation is **heuristic** (marker scan) and is intended to catch obvious regressions.
- Linux: the installed `.desktop` file advertises the expected `MimeType=` list
  (including `x-scheme-handler/<scheme>` entries for deep links from `plugins.deep-link.desktop.schemes`)
  and has an `Exec=`
  placeholder so double-click open passes a path/URL.
  - Note: Parquet (`.parquet`) is not consistently defined in distros’ shared-mime-info
    databases. Formula ships a shared-mime-info definition file in Linux bundles so
    `*.parquet` resolves to `application/vnd.apache.parquet` and the `.desktop` association
    actually matches in file managers.

You can run the same checks locally after building:

```bash
# macOS
app="$(find apps/desktop/src-tauri/target -type d -path '*/release/bundle/macos/*.app' -prune -print -quit)"
plutil -p "$app/Contents/Info.plist" | head -n 200
python3 scripts/ci/verify_macos_bundle_associations.py --info-plist "$app/Contents/Info.plist"

# Linux (.deb)
deb="$(find apps/desktop/src-tauri/target -type f -path '*/release/bundle/deb/*.deb' -print -quit)"
tmpdir="$(mktemp -d)"
dpkg-deb -x "$deb" "$tmpdir"
python3 scripts/ci/verify_linux_desktop_integration.py --package-root "$tmpdir"

# Linux (.rpm)
rpm="$(find apps/desktop/src-tauri/target -type f -path '*/release/bundle/rpm/*.rpm' -print -quit)"
tmpdir_rpm="$(mktemp -d)"
# `rpm` is usually a relative path; run `rpm2cpio` from the current directory and only
# `cd` for the extraction destination.
rpm2cpio "$rpm" | (cd "$tmpdir_rpm" && cpio -idm --quiet --no-absolute-filenames)
python3 scripts/ci/verify_linux_desktop_integration.py --package-root "$tmpdir_rpm"

# Windows
# (Run from a Windows machine/runner; this script uses signtool + Windows Installer COM APIs.)
pwsh -NoProfile -ExecutionPolicy Bypass -File ./scripts/validate-windows-bundles.ps1
```

CI note: the release workflow runs this check on Linux (and, by default, macOS/Windows) **after** the Tauri build step, reusing the
already-built artifacts (`--no-build`). If you need to temporarily skip the check on macOS/Windows (e.g. a hosted-runner regression makes
it flaky), set the GitHub Actions variable `FORMULA_COI_CHECK_ALL_PLATFORMS=0` (or `false`) to keep the Linux check while disabling the
non-Linux ones.

## 6) Installer/bundle artifact size reporting + size gate (tagged releases enforced)

The release workflow reports the size of each generated installer/bundle artifact (DMG / MSI / EXE /
AppImage / DEB / RPM / etc) in the GitHub Actions **step summary**, and **fails tagged releases** if
any artifact exceeds the per-artifact size budget (default: **50 MB**).

For debugging, the workflow also writes a machine-readable JSON report (via the `--json` flag) and
uploads it as a GitHub Actions artifact named `desktop-bundle-size-report-*` (historical name; it
contains **installer artifact** sizes).

To catch regressions before tagging, Linux PRs/main also run `.github/workflows/desktop-bundle-size.yml`,
which builds the Linux AppImage/DEB bundles and reports sizes in the workflow summary (optional gating
via `FORMULA_ENFORCE_BUNDLE_SIZE` / `FORMULA_BUNDLE_SIZE_LIMIT_MB`).

That workflow also runs `python scripts/verify_desktop_binary_stripped.py` after the Tauri build step
to fail early if the produced `formula-desktop` binary is not stripped or if debug/symbol sidecar
files accidentally end up in `**/release/bundle/**`.

Other CI workflows that build desktop installers (including `.github/workflows/desktop-bundle-dry-run.yml`
and `.github/workflows/windows-arm64-smoke.yml`) run the same strip verification script so regressions
are caught before tagging a release.

Note: this is an **installer artifact** budget (DMG/MSI/AppImage/etc), not the **frontend asset
download size** budget (compressed JS/CSS/WASM; see `node scripts/frontend_asset_size_report.mjs`
and `pnpm -C apps/desktop check:bundle-size`).

### Rust binary size controls (Cargo release profile)

The largest contributor under our control is the Rust desktop binary (`formula-desktop`). Size is
primarily controlled by the workspace Cargo release profile in the repo root `Cargo.toml`:

- `strip = "symbols"` – ensures release binaries do not ship with symbol/debug info.
- `lto = "thin"` – enables ThinLTO (often shrinks binaries and improves runtime perf).
- `codegen-units = 1` – improves LTO effectiveness and typically reduces size.

The release workflow also runs `python scripts/verify_desktop_binary_stripped.py` after building to
fail the workflow if the produced desktop binary is not stripped (or if symbol sidecar files like
`.pdb`/`.dSYM` end up in the bundle output directory).

### Crash debug symbols (macOS dSYM / Windows PDB)

For post-release crash symbolication, the release workflow packages platform debug symbols separately:

- **macOS:** `*.dSYM.zip` archives (zipped with `ditto` to preserve bundle structure)
- **Windows:** `*.pdb.zip` archives

These are uploaded as **workflow run artifacts** (so they are downloadable by maintainers without
bloating installer bundles).

Note: the release workflow explicitly enables full Cargo debuginfo (`CARGO_PROFILE_RELEASE_DEBUG=2`)
on macOS/Windows so symbol archives contain enough information for symbolication, while shipped
installers remain stripped.

Optional: to also attach symbol archives to the **draft GitHub Release** (for longer-term retention),
set the GitHub Actions repository variable:

- `FORMULA_UPLOAD_DEBUG_SYMBOLS_TO_RELEASE=1` (or `true`)

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
    . scripts/agent-init.sh

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

MSI upgrades/downgrades also rely on a stable **WiX upgrade code**. This repo pins
`bundle.windows.wix.upgradeCode` in `tauri.conf.json` so that future changes to `productName` do not
accidentally break upgrades/downgrades (or create side-by-side installs). Do **not** change the
upgrade code after shipping a release.

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
