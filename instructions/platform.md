# Workstream F: Platform

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Build the native desktop application shell using **Tauri**. Handle system integration, distribution, auto-updates, and platform-specific features.

**The goal:** 10x smaller than Electron, 4-8x less memory, <500ms startup.

---

## Scope

### Your Code

| Location | Purpose |
|----------|---------|
| `apps/desktop/src-tauri/` | Rust backend, Tauri commands |
| `apps/desktop/src-tauri/capabilities/` | Tauri v2 capability allowlists (permissions for core APIs/plugins) |
| `apps/desktop/src-tauri/gen/schemas/` | Generated JSON schemas used to validate Tauri config/capabilities (reference only) |
| `apps/desktop/src/tauri/` | TypeScript Tauri bindings |
| `apps/desktop/` | Desktop app entrypoint, Vite config |
| `apps/web/` | Web build target (secondary) |

### Your Documentation

- **Primary:** [`docs/11-desktop-shell.md`](../docs/11-desktop-shell.md) — Tauri integration, native features, distribution

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  DESKTOP APPLICATION                                            │
├─────────────────────────────────────────────────────────────────┤
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  WEBVIEW (System WebView)                                  │  │
│  │  ├── React UI Components                                   │  │
│  │  ├── Canvas Grid Renderer                                  │  │
│  │  └── TypeScript Application Logic                          │  │
│  └─────────────────────────┬─────────────────────────────────┘  │
│                            │ Tauri IPC                          │
│  ┌─────────────────────────▼─────────────────────────────────┐  │
│  │  RUST BACKEND                                              │  │
│  │  ├── Calculation Engine (WASM in Worker)                   │  │
│  │  ├── File I/O (async)                                      │  │
│  │  ├── SQLite Database                                       │  │
│  │  ├── System Integration                                    │  │
│  │  └── Native Dialogs                                        │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Requirements

### System Integration

1. **File associations:** Double-click `.xlsx` opens Formula
2. **Native dialogs:** Open/save file dialogs
3. **Clipboard:** Rich clipboard support (HTML, RTF, images)
   - Frontend entry point: `apps/desktop/src/clipboard/platform/provider.js`
   - Desktop prefers custom Rust IPC commands `clipboard_read` / `clipboard_write` (multi-format),
     with a plain-text-only fallback for oversized text: `clipboard_write_text`,
       with fallbacks to:
       - legacy IPC command names: `read_clipboard` / `write_clipboard` (older builds / main-thread bridging on macOS)
       - `navigator.clipboard` (Web Clipboard API)
      - Note: the legacy Tauri clipboard-manager plugin API (`globalThis.__TAURI__.clipboard.readText` / `writeText`)
        is intentionally not enabled in hardened builds to avoid unbounded IPC payloads.
   - Supported formats: `text/plain`, `text/html`, `text/rtf`, `image/png`.
   - JS-facing image API uses `imagePng: Uint8Array` (raw bytes); over Tauri IPC, PNG is transported as `pngBase64`
      (raw base64, no `data:image/png;base64,` prefix).
4. **Drag and drop:** Files and data
5. **System tray:** Background sync indicator
6. **Global shortcuts:** Capture shortcuts even when unfocused
7. **Notifications:** Native system notifications
8. **Deep links (`formula://...`) / OAuth redirects:** allow OAuth PKCE redirects to round-trip back into the running desktop app
   - The desktop shell registers the `formula://` scheme at runtime via `tauri-plugin-deep-link` (best-effort; see `apps/desktop/src-tauri/src/main.rs`).
   - Redirects are forwarded to the frontend via the `oauth-redirect` event.

### Tauri v2 Capabilities & Permissions

This repo is on **Tauri v2**. Config lives in `apps/desktop/src-tauri/tauri.conf.json`.

Key config fields you'll touch most often:

- `build.devUrl`: URL the desktop WebView loads in dev (Vite server)
- `build.frontendDist`: path to built frontend assets for production builds
- `app.security.headers`: COOP/COEP headers (required for `crossOriginIsolated` / `SharedArrayBuffer`)
- `app.security.csp`: Content Security Policy for the desktop WebView
- `apps/desktop/src-tauri/capabilities/*.json`: explicit IPC permission allowlists (scoped to specific windows/webviews)
- `plugins.*`: plugin configuration (e.g. updater)

`app.security.headers` is especially important for the desktop app because the Pyodide-based Python runtime
prefers running in a Worker with `SharedArrayBuffer` (requires `crossOriginIsolated === true`).
See `docs/11-desktop-shell.md` for details.

Tauri v2 permissions are granted via **capabilities**:

- `apps/desktop/src-tauri/capabilities/*.json` define explicit permission allowlists and scope them to window labels via
  the capability file’s `"windows": [...]` list (matches `app.windows[].label` in `tauri.conf.json`).
  - `apps/desktop/src-tauri/capabilities/main.json` scopes itself to `"windows": ["main"]`.
- Some toolchains also support window-level opt-in via `app.windows[].capabilities` in `apps/desktop/src-tauri/tauri.conf.json`.
  - When present, the main window label is `main` and opts into the `main` capability via `"capabilities": ["main"]`.

Note: window-level `app.windows[].capabilities` is **not supported** by the current tauri-build toolchain in this repo
(guardrailed by `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs`). Capability scoping must be done via the
capability file’s `"windows": [...]` list.

Example excerpt (see `apps/desktop/src-tauri/capabilities/main.json` for the full allowlists):

```jsonc
{
  "$schema": "../gen/schemas/desktop-schema.json",
  "identifier": "main",
  "description": "Permissions for the main desktop window (explicit IPC allowlist).",
  "local": true,
  "windows": ["main"],
  "permissions": [
    "allow-invoke",
    {
      "identifier": "core:allow-invoke",
      "allow": [
        { "command": "network_fetch" }
        // ... (see `apps/desktop/src-tauri/capabilities/main.json` for the full list)
      ]
    },
    { "identifier": "core:event:allow-listen", "allow": [{ "event": "open-file" }] },
    { "identifier": "core:event:allow-emit", "allow": [{ "event": "open-file-ready" }] },
    "core:event:allow-unlisten",
    "dialog:allow-open",
    "dialog:allow-save",
    "dialog:allow-confirm",
     "dialog:allow-message",
     "core:window:allow-hide",
     "core:window:allow-show",
     "core:window:allow-set-focus",
     "core:window:allow-close",
     "updater:allow-check",
     "updater:allow-download",
     "updater:allow-install"
   ]
 }
```

Note: `core:event:allow-unlisten` is granted so the frontend can unregister event listeners it previously installed (to avoid
leaking listeners for one-shot flows).

Note: custom Rust `#[tauri::command]` IPC calls are allowlisted at two layers:

- **Application permission**: `apps/desktop/src-tauri/permissions/allow-invoke.json` defines `allow-invoke` (global command
  list; should match the backend `generate_handler![...]` list and frontend `invoke("...")` usage).
  - Granted to the `main` window by including `"allow-invoke"` in `apps/desktop/src-tauri/capabilities/main.json`’s
    `"permissions"` list.
- **Core permission**: `apps/desktop/src-tauri/capabilities/main.json` also includes a `core:allow-invoke` **object**
  allowlist (`allow: [{ "command": "..." }]`) for per-command allowlisting directly in the capability file.
  - The string form `"core:allow-invoke"` is **never** granted (it enables the default/unscoped allowlist).
  - Keep it explicit (no wildcards) and in sync with `permissions/allow-invoke.json` + frontend `invoke("...")` usage.

Note: external URL opening should go through the `open_external_url` Rust command (scheme allowlist enforced in Rust,
and restricted to the main window + trusted app-local origins) rather than granting the webview direct access to the
shell plugin (`shell:allow-open`).

Note: the desktop app intentionally does **not** grant the clipboard-manager plugin permission surface.
Clipboard reads/writes go through custom Rust commands (`clipboard_read` / `clipboard_write` / `clipboard_write_text`)
which enforce trusted-origin + window checks and apply resource limits during deserialization.

### Validating permission identifiers against the installed Tauri toolchain

Tauri’s permission identifiers are derived from the **exact** versions of Tauri core + enabled plugins in your build
toolchain. To validate that `src-tauri/capabilities/*.json` files only use real, supported identifiers, you can
generate the capability JSON schema and list the available permissions:

```bash
# Generates `apps/desktop/src-tauri/gen/schemas/desktop-schema.json` (ignored by git).
# On Linux this requires the system WebView toolchain (gtk/webkit2gtk) because it compiles the desktop feature set.
bash scripts/cargo_agent.sh check -p desktop --features desktop --lib

# Lists all permission identifiers available to this app (core + enabled plugins).
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls
```

Note: on Tauri v2.9, core permissions use the `core:` prefix (e.g. `core:event:allow-listen`, `core:window:allow-hide`).

CI note: capability permission identifiers are validated against the pinned toolchain by
`node scripts/check-tauri-permissions.mjs` (or `pnpm -C apps/desktop check:tauri-permissions`).

If you add new desktop IPC surface area, you must update the capability allowlists:

- new frontend↔backend events → `core:event:allow-listen` / `core:event:allow-emit`
- new plugin API usage → add the corresponding `*:allow-*` permission string(s)
- new invoked app command (`#[tauri::command]`):
  - register it in `apps/desktop/src-tauri/src/main.rs`
  - add it to `apps/desktop/src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission)
  - add it to the `core:allow-invoke` allowlist in `apps/desktop/src-tauri/capabilities/main.json` (`allow: [{ "command": "..." }]`)
  - enforce window/origin/scope checks in Rust (never trust the webview)

We keep guardrail tests to ensure we don't accidentally broaden the desktop IPC surface:

- **Event allowlists**: enforce the **exact** `core:event:allow-listen` / `core:event:allow-emit` sets (no wildcard / allow-all):
  - `apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts`
- **Core/plugin + invoke permissions**: ensure required plugin permissions are explicitly granted (dialogs, window ops,
  updater, etc), we don't accidentally grant dangerous extras, and `allow-invoke.json`
  stays scoped/explicit and in sync with frontend invoke usage; we never grant the unscoped string form
  `"core:allow-invoke"`, and if `core:allow-invoke` is present it uses the object form with an explicit per-command
  allowlist.
  - `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`
- **App command allowlist**: ensure invokable `#[tauri::command]` surface stays explicit + in sync:
  - `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs`

Filesystem access for Power Query is handled via **custom Rust commands** (e.g. `read_text_file`, `list_dir`)
instead of the optional Tauri FS plugin. Those commands enforce an explicit scope:

- `$HOME/**`
- `$DOCUMENT/**`
- `$DOWNLOADS/**` (if the OS/user has a Downloads dir configured and it exists/canonicalizes successfully; on Linux this may be outside `$HOME` via XDG user dirs)

The scope check uses canonicalization to normalize paths and prevent symlink-based escapes. Privileged commands also enforce
**main-window + trusted app origin** checks via `apps/desktop/src-tauri/src/ipc_origin.rs` (defense-in-depth).

### Auto-Update

- Check for updates on startup
- Background download
- User approval before install
- Updater manifest signature verification (CI enforced)
  - Tagged releases must upload a `latest.json` updater manifest **and** a matching `latest.json.sig`.
  - Release CI verifies `latest.json.sig` matches `latest.json` using the updater public key embedded in
    `apps/desktop/src-tauri/tauri.conf.json → plugins.updater.pubkey` (see
    `scripts/ci/verify-updater-manifest-signature.mjs`).
  - If the signature does not verify (or the pubkey is missing/placeholder), the release workflow fails.
- Rollback capability
  - Tauri does not provide an automatic “revert to previous version” after a successful upgrade.
  - Formula supports a clear **manual downgrade path** via the GitHub Releases page (in-app via the
     updater dialog’s “Open release page” / “Download manually” action).
  - Windows note: the rollback path relies on Windows installers supporting downgrades (installing an
    older version over a newer one). This repo enforces that via
    `apps/desktop/src-tauri/tauri.conf.json -> bundle.windows.allowDowngrades: true`
    (guardrailed by `scripts/ci/check-windows-allow-downgrades.mjs`).
  - Rollback depends on keeping older release assets available (don’t delete prior releases). See
    `docs/11-desktop-shell.md` and `docs/release.md`.

### Distribution

- **macOS:** `.dmg`, notarized and stapled
- **Windows:** `.msi` and `.exe` installers, signed
- **Linux:** `.AppImage`, `.deb`, `.rpm`

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Cold start | <500ms to window visible |
| Time to interactive | <1 second |
| Memory (idle) | <100MB |
| Desktop installer artifact size (DMG/MSI/EXE/AppImage) | <50MB per artifact |
| Frontend asset download size (compressed JS/CSS/WASM) | <10MB total |

**What we measure / enforce**

- **Desktop installer artifacts** (DMG/MSI/EXE/AppImage/etc):
  - Measured by `python scripts/desktop_bundle_size_report.py` (scans `<target>/release/bundle` and `<target>/<triple>/release/bundle` after `tauri build`).
  - Guardrail: `python scripts/verify_desktop_binary_stripped.py` fails CI if the produced `formula-desktop` binary is not stripped
    (or if debug/symbol sidecar files like `.pdb`/`.dSYM` accidentally end up in `**/release/bundle/**`).
  - CI gate (tagged releases) in `.github/workflows/release.yml`: **enforced by default**; override via GitHub Actions variables
    `FORMULA_ENFORCE_BUNDLE_SIZE=0` and/or `FORMULA_BUNDLE_SIZE_LIMIT_MB=50`.
  - CI report (Linux PRs + main) in `.github/workflows/desktop-bundle-size.yml` (workflow name: “Desktop installer artifact sizes”):
    informational by default, with optional gating via
    `FORMULA_ENFORCE_BUNDLE_SIZE=1` and `FORMULA_BUNDLE_SIZE_LIMIT_MB=50`. Uploads a JSON report artifact for debugging.
- **Frontend asset download size** (the WebView payload; built Vite `dist/`):
  - Measured by `node scripts/frontend_asset_size_report.mjs --dist apps/desktop/dist` (Brotli total of `dist/assets/**/*.{js,css,wasm}` by default; gzip optional via `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=gzip`).
  - Optional CI gate in `.github/workflows/ci.yml` via GitHub Actions variables:
    - `FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1`
    - `FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB=10`
    - `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=brotli|gzip` (default: brotli)
  - Additional guardrail: `pnpm -C apps/desktop check:bundle-size` enforces tight JS bundle budgets (uncompressed KiB; reports gzip sizes; see `apps/desktop/scripts/bundle-size-check.mjs`).
  - Related (not the 10MB download metric): CI also publishes optional “desktop size” reports that can be gated via GitHub Actions variables:
    - `node scripts/desktop_dist_asset_report.mjs` → `FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB`, `FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB`
    - `python scripts/desktop_size_report.py` → `FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB`, `FORMULA_DESKTOP_DIST_SIZE_LIMIT_MB`

---

## Build & Run

### Development

```bash
# Initialize safe defaults (required for agents)
. scripts/agent-init.sh

# Install dependencies
pnpm install

# Build WASM engine first
pnpm build:wasm
```

Run the desktop frontend (Vite):

```bash
# Vite dev server (matches `build.devUrl` in `apps/desktop/src-tauri/tauri.conf.json`)
pnpm dev:desktop
```

Run the native desktop shell (Tauri):

```bash
# ALWAYS use the cargo wrapper (see `AGENTS.md`)
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri dev
```

> Tip: depending on `build.beforeDevCommand`, `cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri dev` may start Vite for you.
> Avoid running two dev servers on the same port.

### Production Build

```bash
# Initialize safe defaults (required for agents)
. scripts/agent-init.sh

# Build web assets
pnpm build:desktop
```

Build the native app:

```bash
# ALWAYS use the cargo wrapper (see `AGENTS.md`)
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build
```

### Headless Testing

```bash
# Tauri can build without display
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build

# Dev server needs virtual display
cd apps/desktop && xvfb-run --auto-servernum bash ../../scripts/cargo_agent.sh tauri dev
```

---

## Tauri Commands (IPC)

TypeScript ↔ Rust communication:

```rust
// Rust side (e.g. `apps/desktop/src-tauri/src/commands.rs`)
//
// SECURITY: never trust the webview. Validate inputs and enforce authorization/scope
// in Rust before touching filesystem/network/etc.
//
// In this repo, commands must also be:
//  1) registered in `apps/desktop/src-tauri/src/main.rs` (`generate_handler![...]`)
//  2) added to the explicit invoke allowlist in
//     `apps/desktop/src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission)
//
// Note: the capability system gates built-in core/plugin APIs (event/window/dialog/etc) and scopes permissions per-window.
//
// Capabilities are always scoped in the capability file itself via `"windows": [...]` (window labels). Some toolchains also
// support a window-level opt-in layer via `tauri.conf.json` (`app.windows[].capabilities`); when present, keep it in sync
// with the capability file’s `"windows"` list so new windows are unprivileged by default.
//
// App-defined `#[tauri::command]` invocation is allowlisted via `allow-invoke`
// (`apps/desktop/src-tauri/permissions/allow-invoke.json`) and granted to the `main` window via
// `apps/desktop/src-tauri/capabilities/main.json` (the `"allow-invoke"` entry in `"permissions"`).
//
// Note: never grant the string form `"core:allow-invoke"` (default/unscoped allowlist). When `core:allow-invoke` is
// present, it must use the object form with an explicit per-command allowlist (`allow: [{ "command": "..." }]`) and stay
// in sync with actual frontend invoke usage.
//
// Even with allowlisting, commands must be hardened in Rust (trusted-origin + window-label checks, argument validation,
// filesystem/network scope checks, etc).
//
#[tauri::command]
fn check_for_updates(app: tauri::AppHandle, source: crate::updater::UpdateCheckSource) -> Result<(), String> {
    crate::updater::spawn_update_check(&app, source);
    Ok(())
}
```

```typescript
// TypeScript side (desktop renderer, e.g. `apps/desktop/src/*`)
import { getTauriInvokeOrThrow } from "./tauri/api";

const invoke = getTauriInvokeOrThrow();

await invoke("check_for_updates", { source: "manual" });
```

---

## Coordination Points

- **UI Team:** Window management, native dialogs, menu bar
- **File I/O Team:** Native file system access
- **Collaboration Team:** Background sync, system tray status

---

## Platform-Specific Notes

### macOS

- macOS code signing uses `apps/desktop/src-tauri/entitlements.plist` (wired via `bundle.macOS.entitlements` in `tauri.conf.json`).
  - For Developer ID distribution with the hardened runtime, this file must include the WKWebView/JavaScriptCore JIT entitlements (`com.apple.security.cs.allow-jit`, `com.apple.security.cs.allow-unsigned-executable-memory`) or the signed app may launch with a blank WebView.
  - `com.apple.security.network.client` is included so outbound network access (updater/HTTPS) keeps working if the App Sandbox is ever enabled.
  - If we ever enable `com.apple.security.app-sandbox`, we will likely also need `com.apple.security.network.server` because the desktop shell runs an OAuth loopback redirect listener.
  - If we ever enable the App Sandbox in the future, this file is also where sandbox entitlements live.
  - Guardrail: `node scripts/check-macos-entitlements.mjs` (also validated in CI/release workflows).
- Notarize for Gatekeeper
- Support Apple Silicon (aarch64) and Intel (x86_64)
- Use `.icns` icon format

### Windows

- Sign with code signing certificate
- Support both x64 and arm64
- Use `.ico` icon format
- Handle UAC elevation if needed
- Windows uses **Microsoft Edge WebView2**. Installers must ensure the Evergreen runtime is present via
  `bundle.windows.webviewInstallMode` in `apps/desktop/src-tauri/tauri.conf.json` (see `docs/release.md`).

### Linux

- AppImage for universal compatibility
- Respect XDG directories
- Handle Wayland and X11

---

## Reference

- Tauri documentation: https://tauri.app/
- Tauri v2: https://tauri.app/v2/
- wry (WebView): https://github.com/nicklockwood/wry
