# Desktop release process

This repository ships the desktop app via GitHub Releases and Tauri's built-in updater.
Tagged pushes (`vX.Y.Z`) trigger a GitHub Actions workflow that builds installers/bundles for
macOS/Windows/Linux and uploads them to a **draft** GitHub Release.

## Preflight validations (CI enforced)

The release workflow runs a couple of lightweight preflight scripts before it spends time building
bundles. These checks will fail the release workflow on a tagged push if the repo is not in a
releasable state.

Run them locally from the repo root:

```bash
# Ensures the tag version matches apps/desktop/src-tauri/tauri.conf.json "version".
node scripts/check-desktop-version.mjs vX.Y.Z

# Ensures plugins.updater.pubkey/endpoints are not placeholders when the updater is active.
node scripts/check-updater-config.mjs

# Ensures macOS hardened-runtime entitlements include the WKWebView JIT keys
# required for JavaScript/WebAssembly to run in signed/notarized builds.
node scripts/check-macos-entitlements.mjs
```

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

## 2) Tauri updater keys (required for auto-update)

Tauri's updater verifies update artifacts using an Ed25519 signature.

### Generate a keypair

Run this from the repo root (requires the Tauri CLI, `cargo-tauri`). In agent environments, use
the repo cargo wrapper (`scripts/cargo_agent.sh`) instead of bare `cargo`:

```bash
# (Agents) Initialize safe defaults (memory limits, isolated CARGO_HOME, etc.)
source scripts/agent-init.sh

bash scripts/cargo_agent.sh install cargo-tauri --locked
(cd apps/desktop/src-tauri && bash ../../../scripts/cargo_agent.sh tauri signer generate)
```

This prints:
- a **public key** (safe to commit)
- a **private key** (must be stored as a secret)

### Configure the public key in the app

Update `apps/desktop/src-tauri/tauri.conf.json`:

- `plugins.updater.pubkey` → paste the public key (base64 string)
- `plugins.updater.endpoints` → point at your update JSON endpoint(s)

CI note: tagged releases will fail if `plugins.updater.pubkey` or `plugins.updater.endpoints` are
still set to placeholder values. Verify locally with:

```bash
node scripts/check-updater-config.mjs
```

Note: the desktop Rust binary is built with the Cargo feature `desktop` (configured in
`build.features` inside `tauri.conf.json`) so that unit tests can run without the system WebView
toolchain.

### Store the private key in GitHub Actions

Add the following repository secrets:

- `TAURI_PRIVATE_KEY` – the private key string printed by `tauri signer generate` (see above)
- `TAURI_KEY_PASSWORD` – the password used to encrypt the private key (if prompted)

The release workflow passes these to `tauri-apps/tauri-action`, which signs the update artifacts.

## 3) Code signing (optional but recommended)

The release workflow is wired for code signing if the following secrets are present.

### macOS (Developer ID + notarization)

Secrets used by `tauri-apps/tauri-action`:

- `APPLE_CERTIFICATE` – base64-encoded `.p12` Developer ID certificate
- `APPLE_CERTIFICATE_PASSWORD`
- `APPLE_SIGNING_IDENTITY` – e.g. `Developer ID Application: Your Company (TEAMID)`
- `APPLE_ID` – Apple ID email
- `APPLE_PASSWORD` – app-specific password
- `APPLE_TEAM_ID`

#### Hardened runtime entitlements (WKWebView / WASM)

The macOS app is signed with the **hardened runtime**. WKWebView (Tauri/Wry) needs explicit JIT entitlements in the signed binary so that JavaScript and WebAssembly can execute reliably.

The entitlements file used during signing is:

- `apps/desktop/src-tauri/entitlements.plist` (wired via `bundle.macOS.entitlements` in `apps/desktop/src-tauri/tauri.conf.json`)

CI guardrail (run on macOS release builds):

```bash
node scripts/check-macos-entitlements.mjs
```

If these entitlements are missing, a notarized build can still pass notarization but launch with a **blank window** or a crashing WebView process.

#### Local verification checklist (signed app)

1. Build the production bundles:

   ```bash
   pnpm install
   pnpm build:desktop
   cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build
   ```

2. Locate the `.app` produced by Tauri (path can vary by target):

   ```bash
   ls apps/desktop/src-tauri/target/release/bundle/macos/*.app
   ```

3. Verify the signature + entitlements (replace the path as needed):

   ```bash
   app="apps/desktop/src-tauri/target/release/bundle/macos/Formula.app"
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

## 4) Hosting updater endpoints

The desktop app is configured to use **GitHub Releases** as the updater source.

`apps/desktop/src-tauri/tauri.conf.json` points at the `latest.json` manifest generated and uploaded
by `tauri-apps/tauri-action@v0`:

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
    apt-get install -y /deb/*.deb
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

Note: showing a tray icon also requires a desktop environment with **StatusNotifier/AppIndicator**
support (e.g. the GNOME Shell “AppIndicator and KStatusNotifierItem Support” extension).

## Linux: `.rpm` runtime dependencies (Fedora/RHEL)

For RPM-based distros (Fedora/RHEL/CentOS derivatives), the same GTK3/WebKitGTK/AppIndicator stack
must be present at runtime.

These dependencies are declared in `apps/desktop/src-tauri/tauri.conf.json` under
`bundle.linux.rpm.depends` (Fedora/RHEL package names):

- `webkit2gtk4.1` – WebKitGTK system WebView used by Tauri on Linux.
- `gtk3` – GTK3 (windowing/event loop; also required by WebKitGTK).
- `libappindicator-gtk3` – tray icon backend.
- `librsvg2` – SVG rendering used by parts of the GTK icon stack / common icon themes.
- `openssl-libs` – OpenSSL runtime libraries required by native dependencies in the Tauri stack.

### Validating the `.rpm`

After building via `(cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)` (or after CI produces an artifact),
verify the `Requires:` list and shared library resolution.

From `apps/desktop/src-tauri`:

```bash
# Inspect declared dependencies (check webkit2gtk/gtk3/appindicator/etc)
rpm_pkg="$(ls target/release/bundle/rpm/*.rpm | head -n 1)"
rpm -qpR "$rpm_pkg"

# Extract and confirm all linked shared libraries resolve
tmpdir="$(mktemp -d)"
(cd "$tmpdir" && rpm2cpio "$rpm_pkg" | cpio -idmv)
ldd "$tmpdir/usr/bin/formula-desktop" | grep -q "not found" && exit 1 || true
```

For a clean install test (no GUI required), use a Fedora container:

```bash
docker run --rm -it \
  -v "$PWD/target/release/bundle/rpm:/rpm" \
  fedora:40 bash -lc '
    dnf -y install /rpm/*.rpm
    ldd /usr/bin/formula-desktop | grep -q "not found" && exit 1 || true
  '
```

## 5) Verifying a release

After the workflow completes:

1. Open the GitHub Release (draft) and confirm:
   - macOS: `.dmg` (and/or `.app.tar.gz`)
   - Windows: installer (NSIS `.exe` and/or `.msi`)
   - Linux: `.AppImage`, `.deb`, and `.rpm`
2. Download/install on each platform.
3. Publish the release to make it visible to users and (if your updater endpoint references
   GitHub) available for auto-update.

Also verify **cross-origin isolation** is enabled in the packaged app (required for `SharedArrayBuffer` and the Pyodide Worker backend):

- From source (recommended preflight): `pnpm -C apps/desktop check:coi`
- Or in an installed build: ensure there is no startup toast complaining about missing cross-origin isolation, and (if you have DevTools)
  confirm `globalThis.crossOriginIsolated === true`.

## 6) Bundle size reporting + (optional) size gate

The release workflow reports the size of each generated installer/bundle (DMG / MSI / EXE /
AppImage / DEB / etc) in the GitHub Actions **step summary**.

There is also an optional size gate (off by default):

- `FORMULA_ENFORCE_BUNDLE_SIZE=1` → fail the workflow if any artifact exceeds the limit
- `FORMULA_BUNDLE_SIZE_LIMIT_MB=50` → override the default **50 MB** per artifact budget

### Run the size check locally

1. Build the desktop bundles for your platform:

   ```bash
   source scripts/agent-init.sh

   bash scripts/cargo_agent.sh install cargo-tauri --locked
   (cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build)
   ```

2. From the repo root, print a bundle size table:

   ```bash
   python scripts/desktop_bundle_size_report.py
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

**Release hygiene requirements**

1. **Do not delete prior GitHub Releases or assets.**
   - Keep at least several older versions available so users can downgrade when needed.
2. If you mirror artifacts to `releases.formula.app` (or another CDN), ensure you **retain older
   installers/bundles** there too.
   - Users may need to roll back even if the app can't start, so the download URLs must work
     without relying on the updater UI.
