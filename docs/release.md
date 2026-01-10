# Desktop release process

This repository ships the desktop app via GitHub Releases and Tauri's built-in updater.
Tagged pushes (`vX.Y.Z`) trigger a GitHub Actions workflow that builds installers/bundles for
macOS/Windows/Linux and uploads them to a **draft** GitHub Release.

## 1) Versioning + tagging

1. Update the desktop app version in `apps/desktop/src-tauri/tauri.conf.json` (`version`).
2. Merge the version bump to `main`.
3. Create and push a tag:

   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```

The workflow in `.github/workflows/release.yml` will run and create/update a **draft** release with
all platform artifacts attached.

## 2) Tauri updater keys (required for auto-update)

Tauri's updater verifies update artifacts using an Ed25519 signature.

### Generate a keypair

Run this from the Tauri crate directory (requires the Tauri CLI via `cargo tauri`):

```bash
cargo install cargo-tauri --locked
cd apps/desktop/src-tauri
cargo tauri signer generate
```

This prints:
- a **public key** (safe to commit)
- a **private key** (must be stored as a secret)

### Configure the public key in the app

Update `apps/desktop/src-tauri/tauri.conf.json`:

- `plugins.updater.pubkey` → paste the public key (base64 string)
- `plugins.updater.endpoints` → point at your update JSON endpoint(s)

Note: the desktop Rust binary is built with the Cargo feature `desktop` (configured in
`build.features` inside `tauri.conf.json`) so that unit tests can run without the system WebView
toolchain.

### Store the private key in GitHub Actions

Add the following repository secrets:

- `TAURI_PRIVATE_KEY` – the private key string printed by `cargo tauri signer generate`
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

### Windows (Authenticode)

Secrets:

- `WINDOWS_CERTIFICATE` – base64-encoded `.pfx`
- `WINDOWS_CERTIFICATE_PASSWORD`

## 4) Hosting updater endpoints

`apps/desktop/src-tauri/tauri.conf.json` currently uses a placeholder updater endpoint:

```
https://releases.formula.app/{{target}}/{{current_version}}
```

Tauri expects each endpoint to return an `update.json`-style payload (see Tauri updater docs)
describing the latest version for a given target, along with download URLs and signatures.

Two common approaches:

1. **GitHub Releases as the update source**
   - Publish the draft release once validated.
   - Serve the generated update JSON and signature directly from release assets.

2. **Custom update server / CDN**
   - Mirror the update JSON and artifacts from GitHub Releases to `releases.formula.app`.
   - Keep the URL scheme stable so older clients can always resolve update metadata.

## 5) Verifying a release

After the workflow completes:

1. Open the GitHub Release (draft) and confirm:
   - macOS: `.dmg` (and/or `.app.tar.gz`)
   - Windows: installer (NSIS `.exe` and/or `.msi`)
   - Linux: `.AppImage` and/or `.deb`
2. Download/install on each platform.
3. Publish the release to make it visible to users and (if your updater endpoint references
   GitHub) available for auto-update.
