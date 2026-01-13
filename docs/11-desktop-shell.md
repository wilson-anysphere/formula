# Desktop Shell (Tauri v2)

The desktop app is a **Tauri v2.9** shell around the standard web UI. The goal of the Tauri layer is to:

- host the Vite-built UI in a system WebView
- provide native integrations (tray, app menu, global shortcuts, drag/drop + file associations, auto-update)
- expose a small, explicit Rust IPC surface for privileged operations

This document is a “what’s real in the repo” reference for contributors.

## Where the desktop code lives

- **Frontend (TypeScript/Vite):** `apps/desktop/src/`
  - Entry point + desktop host wiring: `apps/desktop/src/main.ts`
  - Desktop wrappers (events, notifications, updater UI, etc): `apps/desktop/src/tauri/`
    - Updater dialog + event handling: `apps/desktop/src/tauri/updaterUi.ts`
    - Notifications wrapper: `apps/desktop/src/tauri/notifications.ts`
    - Startup timings listeners: `apps/desktop/src/tauri/startupMetrics.ts`
    - Open-file IPC helper (`open-file` / `open-file-ready`): `apps/desktop/src/tauri/openFileIpc.ts`
  - Clipboard provider + serialization helpers: `apps/desktop/src/clipboard/`
- **Tauri (Rust):** `apps/desktop/src-tauri/`
  - Tauri config: `apps/desktop/src-tauri/tauri.conf.json`
  - Capabilities (permissions): `apps/desktop/src-tauri/capabilities/main.json`
  - Entry point: `apps/desktop/src-tauri/src/main.rs`
  - IPC commands: `apps/desktop/src-tauri/src/commands.rs`
  - Clipboard commands + platform implementations: `apps/desktop/src-tauri/src/clipboard/`
  - “Open file” path normalization: `apps/desktop/src-tauri/src/open_file.rs`
  - “Open file” IPC queue/handshake state machine: `apps/desktop/src-tauri/src/open_file_ipc.rs`
  - Filesystem scope helpers: `apps/desktop/src-tauri/src/fs_scope.rs` (canonicalization + scope enforcement for all path-taking IPC commands)
  - Custom `asset:` protocol handler (COEP/CORP): `apps/desktop/src-tauri/src/asset_protocol.rs`
  - Stable webview origin helper (used for `asset:` CORS hardening): `apps/desktop/src-tauri/src/tauri_origin.rs`
  - Tray: `apps/desktop/src-tauri/src/tray.rs`
  - Tray status (icon + tooltip updates): `apps/desktop/src-tauri/src/tray_status.rs`
  - App menu: `apps/desktop/src-tauri/src/menu.rs`
  - Global shortcuts: `apps/desktop/src-tauri/src/shortcuts.rs`
  - Updater integration: `apps/desktop/src-tauri/src/updater.rs`

## Startup performance instrumentation

The desktop shell reports real startup timings from the Rust host + webview so we can track cold-start regressions.

Events emitted by the Rust host (to the `main` window):

- `startup:window-visible` — `number` (milliseconds since native process start)
- `startup:webview-loaded` — `number` (milliseconds since native process start)
- `startup:first-render` — `number` (milliseconds since native process start; grid visible)
- `startup:tti` — `number` (milliseconds since native process start; time-to-interactive)
- `startup:metrics` — snapshot payload containing some/all of `{ window_visible_ms, webview_loaded_ms, first_render_ms, tti_ms }`

The frontend installs listeners in `apps/desktop/src/tauri/startupMetrics.ts` and mirrors the latest snapshot into
`globalThis.__FORMULA_STARTUP_TIMINGS__`.

### Viewing startup timings

- **Dev builds**: the Rust host prints a single line to stdout once TTI is reported, e.g.
  ```
  [startup] window_visible_ms=123 webview_loaded_ms=234 first_render_ms=345 tti_ms=456
  ```
- **Release builds**: set `FORMULA_STARTUP_METRICS=1` to enable the same log line.
- **Frontend access**: inspect `globalThis.__FORMULA_STARTUP_TIMINGS__` in DevTools.
- **Multi-run benchmark (recommended)**: from the repo root, use:
  ```bash
  pnpm perf:desktop-startup
  ```
  This command:
  - builds `apps/desktop/dist` (Vite)
  - builds `target/release/formula-desktop` (Rust, `--features desktop`)
  - runs `apps/desktop/tests/performance/desktop-startup-runner.ts`
  - uses a repo-local HOME (`target/perf-home`) so runs don't touch your real `~/.config` / `~/Library`
    - override with `FORMULA_PERF_HOME=/path/to/dir`
    - set `FORMULA_PERF_PRESERVE_HOME=1` to reuse the perf HOME between invocations

  Tuning knobs:
  - `FORMULA_DESKTOP_STARTUP_RUNS` (default: 20)
  - `FORMULA_DESKTOP_STARTUP_TIMEOUT_MS` (default: 15000)
  - `FORMULA_DESKTOP_WINDOW_VISIBLE_TARGET_MS` (default: 500)
  - `FORMULA_DESKTOP_FIRST_RENDER_TARGET_MS` (default: 500)
  - `FORMULA_DESKTOP_TTI_TARGET_MS` (default: 1000)
  - `FORMULA_ENFORCE_DESKTOP_STARTUP_BENCH=1` to fail the command when p95 exceeds the targets (useful for CI gating)
  - `FORMULA_RUN_DESKTOP_STARTUP_BENCH=1` to allow running in CI (the runner skips in CI by default)
  - `FORMULA_DESKTOP_BIN=/path/to/formula-desktop` to benchmark a custom binary

  You can also invoke the runner directly:
  ```bash
  node scripts/run-node-ts.mjs apps/desktop/tests/performance/desktop-startup-runner.ts --bin target/release/formula-desktop --runs 20
  ```

### Idle memory benchmark (desktop process RSS)

To measure idle memory for the desktop app (after TTI, with an empty workbook), run:

```bash
pnpm perf:desktop-memory
```

This reports `idleRssMb`, which is the **resident set size (RSS)** of the desktop process *plus its child processes*,
sampled after the app becomes interactive and a short "settle" delay.

The CI performance suite (`pnpm benchmark`) also reports this as a tracked benchmark metric:

- `desktop.memory.idle_rss_mb.p95` (unit: `mb`)

CI uses `FORMULA_DESKTOP_IDLE_RSS_TARGET_MB` as an absolute budget for this metric (default is intentionally conservative;
tighten it as the desktop app’s idle footprint improves).

The perf commands use a repo-local HOME (`target/perf-home`) by default:

- override with `FORMULA_PERF_HOME=/path/to/dir`
- set `FORMULA_PERF_PRESERVE_HOME=1` to reuse the perf HOME between invocations

Tuning knobs:

- `FORMULA_DESKTOP_MEMORY_RUNS` (default: 10)
- `FORMULA_DESKTOP_MEMORY_SETTLE_MS` (default: 5000)
- `FORMULA_DESKTOP_MEMORY_TIMEOUT_MS` (default: 30000)
- `FORMULA_DESKTOP_IDLE_RSS_TARGET_MB` (default: 100) to set a budget
- `FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH=1` (or `--enforce`) to fail the command when p95 exceeds the budget
- `FORMULA_RUN_DESKTOP_MEMORY_BENCH=1` to allow running in CI (the runner skips in CI by default)

### Size report (dist + binary + bundles)

To get a quick size breakdown for the desktop app, run:

```bash
pnpm perf:desktop-size
```

This reports:

- `apps/desktop/dist` total size (and largest assets)
- frontend asset download size (compressed JS/CSS/WASM) via `scripts/frontend_asset_size_report.mjs`
- the built desktop binary size (`target/**/formula-desktop`)
- if present, installer/bundle artifacts under `target/**/release/bundle` (via `scripts/desktop_bundle_size_report.py`)

Installer artifact size gating knobs (used by the release workflow; DMG/MSI/AppImage/etc):

- `FORMULA_BUNDLE_SIZE_LIMIT_MB` (default: 50MB per artifact)
- `FORMULA_ENFORCE_BUNDLE_SIZE=1` to fail when any artifact exceeds the limit

Frontend asset download size gating knobs (compressed JS/CSS/WASM under `dist/assets`):

- `FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB` (default: 10MB total)
- `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=brotli|gzip` (default: brotli)
- `FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1` to fail when the total exceeds the limit

---

## Tauri configuration (v2)

The desktop configuration lives in `apps/desktop/src-tauri/tauri.conf.json` (Tauri v2 format).

Key sections you’ll most commonly touch:

### App identity (name/id/version)

Top-level keys in `tauri.conf.json` define the packaged app identity:

- `productName`: human-readable app name
- `identifier`: reverse-DNS bundle identifier (`app.formula.desktop`)
- `version`: desktop app version used by the updater / release tooling. Tagged releases (`vX.Y.Z`)
  must match this value (CI enforces this via `scripts/check-desktop-version.mjs`; see
  `docs/release.md`).
- `mainBinaryName`: the Rust binary name Tauri expects to launch (matches `[[bin]].name` in `apps/desktop/src-tauri/Cargo.toml`)

### `build.*` (frontend dev/build + Cargo feature flags)

- `build.beforeDevCommand`: `pnpm dev` (runs Vite)
- `build.beforeBuildCommand`: `pnpm build` (builds `../dist`)
- `build.devUrl`: `http://localhost:4174` (matches `apps/desktop/package.json`)
- `build.frontendDist`: `../dist`
- `build.features: ["desktop"]` enables the Cargo feature gate for the real desktop binary (see “Cargo feature gating” below)

### `app.security.headers` / `app.security.csp`

`app.security.headers` configures response headers for the built-in `tauri://…` protocol. In this repo it is used to set COOP/COEP (see “Cross-origin isolation” below).

The CSP is set in `app.security.csp` (see `apps/desktop/src-tauri/tauri.conf.json`).

Current policy (exact):

```text
default-src 'self'; base-uri 'self'; form-action 'self'; navigate-to 'self'; object-src 'none'; frame-ancestors 'none'; img-src 'self' asset: data:; style-src 'self' 'unsafe-inline'; script-src 'self' 'wasm-unsafe-eval' 'unsafe-eval' blob: data:; worker-src 'self' blob: data:; child-src 'self' blob: data:; connect-src 'self' https: ws: wss: blob: data:
```

Rationale:

- `form-action 'self'` prevents accidental/malicious form submissions from the WebView to unexpected origins.
- `navigate-to 'self'` prevents in-webview navigations to remote origins. External links and OAuth flows should be opened
  via the OS browser (see `open_external_url` / `shellOpen`).
- The Rust engine runs as **WebAssembly inside a module Worker**, so CSP must allow:
  - `script-src 'wasm-unsafe-eval'` for WASM compilation/instantiation.
  - `worker-src 'self' blob: data:` for module workers (Vite may use `blob:`/`data:` URLs for worker bootstrapping).
- The extension runtime (`BrowserExtensionHost`) also runs each extension in a **module Worker** loaded from an in-memory
  `blob:`/`data:` module URL, so CSP must allow `worker-src blob:` and `script-src blob: data:`.
- Extension panels are rendered as sandboxed **`blob:` iframes**, so CSP must allow `child-src blob:` (or `frame-src blob:`)
  to avoid blocking the iframe load.
- We also rely on `script-src 'unsafe-eval'` for the scripting sandbox (`new Function`-based evaluation in a Worker).
- `connect-src` is intentionally restrictive (no `http:`), but allows outbound network for collaboration + extensions
  (`ws:`/`wss:`) and HTTPS APIs (`https:`), along with same-origin + `blob:`/`data:` URLs.
  - Note: Rust IPC network (`network_fetch`, `marketplace_*`) is performed by the desktop backend (reqwest) and is not
    governed by the WebView CSP. Those commands currently allow `http:` URLs (useful for local dev servers) in addition to
    `https:`.

### Network strategy (extensions + marketplace)

In packaged desktop builds we keep a restrictive CSP that avoids enabling `http:` and only permits `https:` plus
WebSockets (`ws:`/`wss:`) and app-local (`'self'`) / in-memory (`blob:`/`data:`) URLs.

Network access is mediated at two layers:

- **Extensions:** the extension worker runtime (`packages/extension-host/src/browser/extension-worker.mjs`) hides
  `__TAURI__` and replaces browser networking primitives (`fetch`, `WebSocket`, `XMLHttpRequest`, etc) with
  permission-gated wrappers. This makes the `network` permission checks in `BrowserExtensionHost` the enforcement point
  (not CSP).
  - `formula.network.fetch(...)` (and `fetch(...)` inside extensions) is implemented by:
    - **Tauri desktop**: `invoke("network_fetch", ...)` (Rust/reqwest; avoids CORS and enforces an `http(s)` scheme allowlist).
    - **Web / non-Tauri**: a `fetch(...)` fallback.
- `formula.network.openWebSocket(...)` is a permission check; the actual socket is opened directly in the extension
  worker via `new WebSocket(...)` (hence `ws:`/`wss:` in `connect-src`).
- **Marketplace:** `MarketplaceClient` prefers Rust IPC (`marketplace_search`, `marketplace_get_extension`,
  `marketplace_download_package`) when running under Tauri with an absolute `http(s)` base URL. In other contexts it
  falls back to `fetch(...)`.
- Note: these Rust IPC commands (`network_fetch`, `marketplace_*`) also enforce **main-window + trusted app origin** checks
  via `apps/desktop/src-tauri/src/ipc_origin.rs` (defense-in-depth).

Rust IPC implementations live in `apps/desktop/src-tauri/src/commands.rs`.

### Tauri v2 capabilities (permissions)

Tauri v2 replaces Tauri v1’s “allowlist” with **capabilities**, defined as JSON files under:

- `apps/desktop/src-tauri/capabilities/` (main capability: `capabilities/main.json`)

Capabilities are scoped per window by the capability file’s `"windows": [...]` list (window labels from
`apps/desktop/src-tauri/tauri.conf.json`).

Note: some Tauri toolchains support window-level opt-in via `app.windows[].capabilities`, but the current tauri-build
toolchain used in this repo rejects that field. Keep capability scoping in the capability file itself (guardrailed by
`apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs`).

Example excerpt:

```jsonc
// apps/desktop/src-tauri/capabilities/main.json
{
  "identifier": "main",
  "windows": ["main"],
  "permissions": [
    "allow-invoke",
    // Tauri core permission for invoke scoping. Keep this in object form with an explicit allowlist
    // (never the string form `core:allow-invoke` which behaves like an allow-all/default allowlist).
    {
      "identifier": "core:allow-invoke",
      "allow": [
        { "command": "network_fetch" },
        { "command": "open_external_url" }
        // ... (see the real file for the full list)
      ]
    },
    { "identifier": "core:event:allow-listen", "allow": [{ "event": "open-file" }] },
    { "identifier": "core:event:allow-emit", "allow": [{ "event": "open-file-ready" }] },
    "core:event:allow-unlisten",
    "dialog:allow-open",
    "core:window:allow-set-focus",
    "clipboard-manager:allow-read-text",
    "clipboard-manager:allow-write-text",
    "updater:allow-check"
  ]
}
```

When adding new uses of privileged plugin APIs (clipboard/dialog/updater/window APIs) or adding new desktop event names,
update the relevant allowlists in `capabilities/main.json`.

When adding a new Rust `#[tauri::command]` invoked from the frontend, also update the invoke allowlist in:

- `apps/desktop/src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission; guardrailed by `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs` and `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`)
- `apps/desktop/src-tauri/capabilities/main.json` (`core:allow-invoke` object form allowlist; guardrailed by `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`)

The `allow-invoke` permission is granted to the `main` window via `apps/desktop/src-tauri/capabilities/main.json` by
including `"allow-invoke"` in that capability’s `"permissions"` list.

This is guardrailed by `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`, which ensures the command
allowlist is explicit (no wildcards/duplicates) and matches actual frontend `invoke("...")` usage. It also asserts we do
not grant the unscoped string form `core:allow-invoke` (default allowlist). This repo grants `core:allow-invoke` in the
**object form** `{ "identifier": "core:allow-invoke", "allow": [{ "command": "..." }] }`; keep it explicit and in sync
with `permissions/allow-invoke.json` + actual frontend `invoke("...")` usage.

See “Tauri v2 Capabilities & Permissions” below for the concrete `main.json` contents.

### Cross-origin isolation (COOP/COEP) for Pyodide / `SharedArrayBuffer`

The Pyodide-based Python runtime prefers running in a **Worker** with a `SharedArrayBuffer + Atomics` bridge.
In Chromium/WebView2, that requires a **cross-origin isolated** browsing context:

- `globalThis.crossOriginIsolated === true`
- `typeof SharedArrayBuffer !== "undefined"`

How this is (currently) handled in the repo:

- **Dev / preview (Vite):** `apps/desktop/vite.config.ts` sets COOP/COEP headers on dev/preview responses.
- **Packaged Tauri builds:** COOP/COEP are set via `app.security.headers` in `apps/desktop/src-tauri/tauri.conf.json`,
  which Tauri applies to its built-in `tauri://…` protocol responses.
  - Additionally, the desktop shell overrides the `asset:` protocol handler (see `apps/desktop/src-tauri/src/asset_protocol.rs`)
    to attach `Cross-Origin-Resource-Policy: cross-origin` so `convertFileSrc(...)` URLs can still be embedded when
    `Cross-Origin-Embedder-Policy: require-corp` is enabled.
    - For security, it does **not** set `Access-Control-Allow-Origin: *`; it sets `Access-Control-Allow-Origin` to the
      **stable initial webview origin** (mirroring Tauri’s upstream `window_origin` behavior) so an external navigation
      cannot gain CORS access to arbitrary `asset://…` files. See `apps/desktop/src-tauri/src/tauri_origin.rs`.
    - Security boundary: `asset://...` responses are only served to **trusted app-local origins**
      (`localhost`, `127.0.0.1`, `::1`, `*.localhost`, best-effort `file://`). If the WebView navigates to remote/untrusted
      content, all `asset:` requests are denied with `403` to avoid turning `asset:` into a local-file read primitive.
    - DoS hardening: non-range `asset:` responses are **size-limited** (currently 10 MiB) to avoid unbounded in-memory file
      reads; large files must be accessed via `Range` requests (which are already clamped per-request).
  - If isolation is missing in a production desktop build, the UI logs an error and shows a long-lived toast (see
    `warnIfMissingCrossOriginIsolationInTauriProd()` in `apps/desktop/src/main.ts`).

Quick verification guidance lives in `apps/desktop/README.md` (“Production/Tauri: `crossOriginIsolated` check”),
including an automated smoke check:

```bash
pnpm -C apps/desktop check:coi
```

This check validates `globalThis.crossOriginIsolated`, `SharedArrayBuffer` availability, and that a basic Web Worker can start (to catch CSP / asset-protocol regressions that would break the Pyodide worker backend).

Release CI runs this check on macOS/Windows/Linux before uploading artifacts. If you need to temporarily skip it on macOS/Windows
(e.g. a hosted-runner regression makes it flaky), set the GitHub Actions variable `FORMULA_COI_CHECK_ALL_PLATFORMS=0` (or `false`) to keep the
Linux check while disabling the non-Linux ones.

To validate `asset://` (i.e. `convertFileSrc`) resources still load under COEP, the repo also includes:

- `apps/desktop/asset-protocol-test.html` (open in the desktop app and follow the instructions on the page)

Practical warning: with `Cross-Origin-Embedder-Policy: require-corp`, *every* subresource must be same-origin or explicitly opt-in via CORS/CORP.
In Tauri, `convertFileSrc(...)` produces `asset://...` URLs; those `asset:` responses need a CORP header or they won’t load under COEP.
The repo’s custom `asset:` handler adds `Cross-Origin-Resource-Policy: cross-origin` for this reason.

### `bundle.*` (packaging)

Release CI (`.github/workflows/release.yml`) produces platform installers/bundles for **macOS/Windows/Linux**
(including multi-arch artifacts like macOS universal and Windows x64+arm64). For the expected artifact
list and verification commands, see `docs/release.md` (“Verifying a release”).

Notable keys:

- `bundle.fileAssociations` registers spreadsheet file types with the OS:
  `.xlsx`, `.xls`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`, `.xlsb`, `.csv`, `.parquet`.
  - `.parquet` open support is behind the Cargo `parquet` feature (enabled by the `desktop` feature; see `apps/desktop/src-tauri/Cargo.toml` and `apps/desktop/src-tauri/src/open_file.rs`).
- `bundle.linux.deb.depends` documents runtime deps for Linux packaging (e.g. `libwebkit2gtk-4.1-0`, `libgtk-3-0t64 | libgtk-3-0`,
  appindicator, `librsvg2-2`, `libssl3t64 | libssl3`).
- `bundle.linux.rpm.depends` documents runtime deps for RPM-based distros using **RPM rich dependencies**
  (e.g. `(webkit2gtk4.1 or libwebkit2gtk-4_1-0)`, `(gtk3 or libgtk-3-0)`,
  `((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))`,
  `(librsvg2 or librsvg-2-2)`, `(openssl-libs or libopenssl3)`).
- `bundle.macOS.entitlements` / signing keys and `bundle.windows.timestampUrl`.
  - `bundle.macOS.entitlements` points at `apps/desktop/src-tauri/entitlements.plist`.
    - For Developer ID distribution with the hardened runtime, the entitlements must include the WKWebView JIT keys (`com.apple.security.cs.allow-jit`, `com.apple.security.cs.allow-unsigned-executable-memory`) or the signed app may launch with a blank WebView. See `docs/release.md` for troubleshooting.
    - Guardrail: `node scripts/check-macos-entitlements.mjs` (also run in CI).
  - `bundle.windows.timestampUrl` should be an **HTTPS** timestamp server (avoid plaintext HTTP Authenticode timestamping). Release CI guardrails enforce this (see `scripts/ci/check-windows-timestamp-url.mjs` and `apps/desktop/src/tauri/__tests__/tauriSecurityConfig.vitest.ts`).
- `bundle.windows.webviewInstallMode` controls how Windows installers ensure the Microsoft Edge **WebView2** runtime is present.
  - This repo uses the Evergreen bootstrapper mode so installs work on machines without WebView2 (see `docs/release.md` for details and offline alternatives).

### Distribution (GitHub Releases)

The desktop app is shipped via **GitHub Releases** (see `docs/release.md` for the full checklist).
Tagged builds are expected to produce:

- **macOS:** universal `.dmg` (plus updater payload `.app.tar.gz`)
- **Windows:** installers for **x64** and **ARM64** (`.msi` + `.exe`)
- **Linux:** `.AppImage` + `.deb` + `.rpm`

Auto-update is driven by the Tauri updater manifest (`latest.json`) uploaded to the release.
For the exact `latest.json.platforms` key names (multi-arch), see `docs/desktop-updater-target-mapping.md`.

### `plugins.updater`

Auto-update is configured under `plugins.updater` (Tauri v2 plugin config). Release builds embed an
updater public key (`pubkey`) and fetch update metadata from `endpoints` (this repo defaults to the
GitHub Releases `latest.json` manifest; see `docs/release.md`):

- `plugins.updater.pubkey` → updater public key (base64; safe to commit; committed in this repo). Must match the private key used in CI (`TAURI_PRIVATE_KEY`) to sign update artifacts (see `docs/release.md`).
- `plugins.updater.endpoints` → update JSON endpoint(s). This repo defaults to the GitHub Releases manifest:
  - `https://github.com/wilson-anysphere/formula/releases/latest/download/latest.json`
  - (The matching signature, `latest.json.sig`, is uploaded by `tauri-action` and verified using `pubkey`.)
- `plugins.updater.dialog: false` → the Rust host emits events instead of showing a built-in dialog (custom UI in the frontend)
- `plugins.updater.windows.installMode` controls the Windows update install mode (currently `passive`)

Frontend event contract (emitted by `apps/desktop/src-tauri/src/updater.rs`):

- `update-check-started` – payload: `{ source: "startup" | "manual" }`
- `update-check-already-running` – payload: `{ source: "manual" }`
- `update-available` – payload: `{ source, version, body? }`
- `update-not-available` – payload: `{ source }`
- `update-check-error` – payload: `{ source, message }`
- `update-download-started` – payload: `{ source, version }` (best-effort background download)
- `update-download-progress` – payload: `{ source, version, chunkLength, downloaded, total?, percent? }`
- `update-downloaded` – payload: `{ source, version }`
- `update-download-error` – payload: `{ source, version, message }`

The Rust host guards updater checks with a single in-flight flag so **only one network check runs at a
time**. If a check is already running, additional **manual** triggers emit
`update-check-already-running` so the UI can show “Already checking…”. Additional **startup** triggers
are ignored silently.

When an update is found, the backend also starts a **best-effort background download** so the user can
restart/apply without waiting for a second download. The frontend consumes `update-downloaded` by
showing a lightweight “Update ready to install” toast and uses the `install_downloaded_update` command
as part of the restart-to-install flow (falling back to the updater plugin API if needed).

Release CI note: when `plugins.updater.active=true`, tagged releases validate `pubkey`/`endpoints`
via `node scripts/check-updater-config.mjs`.

### `plugins.notification`

Native system notifications are enabled via `plugins.notification` and are used for lightweight UX
signals (e.g. “Update available”, “Power Query refresh complete”). The frontend wrapper lives in
`apps/desktop/src/tauri/notifications.ts`.

Security notes:

- The Rust command `show_system_notification` (exposed via `__TAURI__.core.invoke(...)`) is restricted to the
  main window and to trusted app-local origins.
- We intentionally do **not** grant notification plugin permissions (e.g. `notification:*` / `core:notification:*`) to the
  webview in `apps/desktop/src-tauri/capabilities/main.json`. System notifications are instead routed through
  `invoke("show_system_notification", ...)` so the main-window + trusted-origin checks are always enforced.
- The frontend `notify(...)` helper tries the Tauri notification plugin API first; if that API is unavailable or blocked by
  permissions (expected in hardened builds), it falls back to `invoke("show_system_notification", ...)`.
- In **web builds** (no `__TAURI__`), `notify(...)` can fall back to the Web Notification API (permission-gated). In **desktop
  builds**, it intentionally does **not** fall back to the Web Notification API to avoid untrusted/navigated-to content
  triggering system notifications outside the hardened Rust command path.

Minimal excerpt (not copy/pasteable; see the full file for everything):

```jsonc
// apps/desktop/src-tauri/tauri.conf.json
{
  "productName": "Formula",
  "mainBinaryName": "formula-desktop",
  "version": "0.1.0",
  "identifier": "app.formula.desktop",
  "build": {
    "beforeBuildCommand": "pnpm build",
    "beforeDevCommand": "pnpm dev",
    "devUrl": "http://localhost:4174",
    "frontendDist": "../dist",
    "features": ["desktop"]
  },
  "app": {
    "security": {
      "headers": {
        "Cross-Origin-Opener-Policy": "same-origin",
        "Cross-Origin-Embedder-Policy": "require-corp"
      },
      "csp": "..." // see `apps/desktop/src-tauri/tauri.conf.json` for the full, current CSP
    },
    "windows": [
      { "label": "main", "title": "Formula", "width": 1280, "height": 800, "dragDropEnabled": true }
    ]
  },
  "bundle": {
    "fileAssociations": [{ "ext": ["xlsx"], "name": "Excel Spreadsheet", "role": "Editor" }]
    // (Other bundle config omitted for brevity; see the real file.)
    //
    // Note: `formula://` deep links are registered at runtime via `tauri-plugin-deep-link`,
    // not via `bundle.protocols`.
  },
  "plugins": {
    "updater": {
      "active": true,
      "dialog": false,
      "endpoints": ["https://github.com/wilson-anysphere/formula/releases/latest/download/latest.json"],
      "pubkey": "<updater public key (see apps/desktop/src-tauri/tauri.conf.json)>"
    },
    "notification": {}
  }
}
```

Note: calling the updater plugin from the **frontend** (via `globalThis.__TAURI__.updater`) is gated by Tauri v2 window
capabilities. If you add/update updater UI flows, ensure the relevant `updater:allow-*` permissions are present in
`apps/desktop/src-tauri/capabilities/main.json` (for example: `updater:allow-check`, `updater:allow-download`, `updater:allow-install`).

#### Rollback / downgrade (rollback capability)

Tauri's updater flow is designed to be **failure-safe** (signature verification + install handoff to the
platform installer). If an update **fails to download or fails to install**, the current version
should remain installed.

Tauri does **not** keep multiple versions installed and does **not** provide a one-click "revert to
previous version" after a successful upgrade.

To satisfy the platform requirement **"Rollback capability"**, Formula supports a clear **manual
downgrade path**:

The updater dialog includes an **"Open release page"** action. If an update download/install fails,
that action is relabeled/promoted to **"Download manually"** and the dialog surfaces manual
download/downgrade instructions.

1. Open Formula's **Releases** page:
   - In-app: **Help → Open Release Page**, or via the updater dialog's **"Open release page"** / **"Download manually"**
     action.
   - Browser: https://github.com/wilson-anysphere/formula/releases
2. Download the installer/bundle for the **older version** you want.
3. Install it over your current install (or uninstall first if your platform's installer blocks
   downgrades).

**Platform notes**

- **Windows (x64 + ARM64, NSIS/MSI):**
  - Download the installer that matches your machine (**x64** vs **ARM64**).
  - Formula's Windows installers are configured to **allow downgrades** (`bundle.windows.allowDowngrades: true`).
  - **NSIS `.exe`**: when downgrading, the installer will show a maintenance screen; for the cleanest rollback choose
    **“Uninstall before installing”**, then proceed.
  - **WiX `.msi`**: if your currently installed Formula version was installed via **MSI** (including installs performed by
    the in-app auto-updater), running an older MSI will remove the installed MSI version and then install the selected version.
  - Tip: prefer using the **same installer format** you originally installed with (`.exe` ↔ `.exe`, or `.msi` ↔ `.msi`).
  - If an installer still refuses to proceed (e.g. “a newer version is already installed”), uninstall Formula from
    *Apps & Features* and then run the older installer.
- **macOS (universal `.dmg`):**
  - The macOS build is universal, so the same `.dmg` works on Intel and Apple Silicon.
  - Download the `.dmg`, open it, then drag `Formula.app` into `/Applications`.
  - macOS will prompt to **Replace** the existing app; confirm to downgrade.
- **Linux (.AppImage / .deb / .rpm):**
  - **AppImage:** download the older AppImage and replace the current file.
  - **deb/rpm:** install the older package with your package manager (some distros require an
    explicit downgrade flag).

Important: rollback depends on old versions staying available. See `docs/release.md` — we must not
delete prior release assets.

---

## Rust host (Tauri backend)

### Entry point: `apps/desktop/src-tauri/src/main.rs`

`main.rs` wires together:

- **state** (`SharedAppState`) + **macro trust store** (`SharedMacroTrustStore`)
- Tauri plugins:
  - `tauri_plugin_global_shortcut` (registers accelerators + emits app events)
  - `tauri_plugin_dialog` (native open/save dialogs; gated by `dialog:allow-*` permissions)
  - `tauri_plugin_clipboard_manager` (plain-text clipboard helpers; gated by `clipboard-manager:allow-*` permissions)
  - `tauri_plugin_shell` (used by the Rust command `open_external_url` to open external links in the host OS; direct webview access is not granted)
  - `tauri_plugin_updater` (update checks)
  - `tauri_plugin_notification` (native notifications)
  - `tauri_plugin_single_instance` (forward argv/cwd from subsequent launches into the running instance)
  - `tauri_plugin_deep_link` (best-effort runtime registration for the `formula://` deep link scheme)
- A custom `asset:` protocol handler (`asset_protocol.rs`) to attach COEP/CORP-friendly headers for `asset://...` URLs (used by `convertFileSrc`).
- App menu setup (see `apps/desktop/src-tauri/src/menu.rs`) and `.on_menu_event(...)` forwarding.
- `invoke_handler(...)` mapping commands in `commands.rs`
- window/tray event forwarding to the frontend via `app.emit(...)` / `window.emit(...)`

#### Close flow (hide vs quit)

The desktop app deliberately **does not exit on window close** so the tray remains available.

The window-close sequence is:

1. Rust receives `WindowEvent::CloseRequested` and calls `api.prevent_close()`.
2. Rust emits `close-prep` with a random token.
3. Frontend (in `apps/desktop/src/main.ts`) commits any in-progress edits, flushes pending workbook sync, calls
   `set_macro_ui_context`, then emits `close-prep-done` with the same token.
4. Rust runs a best-effort `Workbook_BeforeClose` macro (if trusted) and collects any cell updates.
5. Rust emits `close-requested` with `{ token, updates }`.
6. Frontend applies any macro cell updates, prompts for unsaved changes if needed, then either:
   - hides the window (default behavior; app keeps running in the tray), or
   - keeps the window open if the user cancels the close (e.g. cancels the unsaved-changes prompt)
7. Frontend emits `close-handled` with the token so Rust can clear its “close in flight” guard.

Other close entry points (e.g. **menu Close Window** / `Cmd/Ctrl+W`) are handled entirely in the frontend. In those cases, the
frontend runs `Workbook_BeforeClose` as a best-effort (trusted-only, no permission escalation) via the `fire_workbook_before_close`
command, applies any updates, and then follows the same “unsaved changes prompt → hide vs quit” decision.

Implementation detail: `main.rs` uses an `AtomicBool` (`CLOSE_REQUEST_IN_FLIGHT`) to prevent overlapping close flows if the user clicks close repeatedly while a prompt is still open.

#### Drag & drop → open file

When a file is dropped onto the window, `main.rs` listens for `WindowEvent::DragDrop` and emits:

- `file-dropped` with `Vec<String>` of filesystem paths

The frontend listens for this event and queues an open via `queueOpenWorkbook(...)` (so opens are serialized).

#### Open-with / file associations / CLI args

In addition to drag & drop, the desktop shell supports opening workbooks via:

- “Open with…” / Finder / Explorer (file associations configured in `bundle.fileAssociations` in `tauri.conf.json`)
- passing a path on the command line (cold start)
- launching the app again while an instance is already running (warm start)

Implementation notes:

- `apps/desktop/src-tauri/src/open_file.rs` extracts supported spreadsheet paths from argv-style inputs (and also supports `file://...` URLs used by macOS open-document events).
  - Extension-based support is case-insensitive: `xlsx`, `xls`, `xlsm`, `xltx`, `xltm`, `xlam`, `xlsb`, `csv` (+ `parquet` when compiled with the `parquet` feature).
  - If the extension is missing or unsupported, it falls back to a lightweight content sniff (`file_io::looks_like_workbook`) so “downloaded/renamed” workbooks can still be opened from OS open-file events.
- `main.rs` uses a small in-memory queue (`OpenFileState`, implemented in `apps/desktop/src-tauri/src/open_file_ipc.rs`) so open-file requests received *before* the frontend installs its listeners aren’t lost.
  - Backend emits: `open-file` (payload: `string[]` paths)
  - Frontend emits: `open-file-ready` once its `listen("open-file", ...)` handler is installed, which flushes any queued paths.
- When an open-file request is handled, `main.rs` **shows + focuses** the main window before emitting `open-file` so the request is visible to the user.
- On macOS, `tauri::RunEvent::Opened { urls, .. }` is routed through the same pipeline so opening a document in Finder reaches the running instance.

#### OAuth redirects (`formula://…` deep links vs RFC 8252 loopback)

Formula Desktop supports two redirect-capture strategies for OAuth (typically PKCE / auth-code flows):

| Redirect URI in the auth request | How it’s captured | When to use |
|---|---|---|
| `formula://…` (custom scheme deep link, e.g. `formula://oauth/callback`) | OS launches/forwards a `formula://…` URL to the app (via `tauri-plugin-deep-link` + argv/single-instance handling); Rust forwards it to the frontend as `oauth-redirect` | **Preferred** when the provider allows custom schemes (no local port binding) |
| `http://127.0.0.1:<port>/…`, `http://localhost:<port>/…`, or `http://[::1]:<port>/…` (loopback) | Frontend detects a loopback `redirect_uri` query param in the auth URL and calls the Rust command `oauth_loopback_listen` to start a temporary local HTTP listener; the listener forwards the observed redirect as `oauth-redirect` | Fallback for providers that reject custom schemes |

Deep-link scheme config/registration:

- Config: `apps/desktop/src-tauri/tauri.conf.json` → `plugins["deep-link"].desktop.schemes: ["formula"]`
- Runtime: `apps/desktop/src-tauri/src/main.rs` attempts best-effort OS registration via `app.handle().deep_link().register("formula")`.

**How the frontend chooses:** `DesktopOAuthBroker.openAuthUrl(...)` (`apps/desktop/src/power-query/oauthBroker.ts`) inspects the auth URL’s `redirect_uri` query param. If it is a supported loopback URI, it invokes `oauth_loopback_listen` **before** opening the system browser; otherwise it relies on `formula://…` deep-link delivery.

Recommended redirect URIs (used by the desktop Power Query UI; see `apps/desktop/src/panels/data-queries/DataQueriesPanelContainer.tsx`):

- Deep link: `formula://oauth/callback`
- Loopback example (choose an unused port): `http://127.0.0.1:4242/oauth/callback`

##### Supported loopback redirect URIs

The loopback listener implementation (`oauth_loopback_listen` in `apps/desktop/src-tauri/src/main.rs`, parser in
`apps/desktop/src-tauri/src/oauth_loopback.rs`) supports:

- **Scheme:** `http` only (no `https`)
- **Host:** `127.0.0.1`, `localhost`, or `[::1]`
  - For `localhost`, the backend attempts both IPv4 and IPv6 bindings so platform resolver differences don’t break the flow.
- **Port:** required and **non-zero**
- **Path:** any, but the inbound request path must match the configured `redirect_uri` path exactly (mismatches return `404`)
- **Query:** preserved and forwarded to the frontend
- **Fragment (`#…`):** not supported (browsers don’t send URL fragments to loopback HTTP servers)
- **Method:** `GET` only
- **Lifetime:** listener stops after ~5 minutes (best-effort timeout)

##### Redirect forwarding to the frontend (`oauth-redirect` + readiness handshake)

Both deep-link and loopback flows end up as the same desktop event:

- **Rust → frontend:** `oauth-redirect` (payload: full redirect URL string)
- **Frontend → Rust:** `oauth-redirect-ready` once `listen("oauth-redirect", ...)` is installed (flushes any queued early redirects)

The backend buffers early redirects in memory (`OauthRedirectState` in `apps/desktop/src-tauri/src/oauth_redirect_ipc.rs`) so fast redirects at cold start aren’t dropped. The frontend listener lives in `apps/desktop/src/main.ts` and forwards URLs into the in-process OAuth broker (`oauthBroker.observeRedirect(...)`).

##### Troubleshooting

- **Provider rejects `formula://…` redirect URIs:** use a loopback redirect and register one of `http://127.0.0.1:<port>/<path>`, `http://localhost:<port>/<path>`, or `http://[::1]:<port>/<path>` with the provider.
- **Provider rejects a specific loopback host:** some providers only allow `localhost` (and not `127.0.0.1` or `::1`) in their allowlist UI. Pick the loopback host form your provider supports.
- **Provider requires an `https://` redirect URI:** loopback capture is `http://` only; use a custom-scheme deep link or a different auth approach if `http://` loopback redirects are disallowed.
- **Deep link doesn’t trigger / app isn’t opened by `formula://…`:** verify the OS has a protocol handler registered for `formula://`. The desktop host attempts best-effort runtime registration on startup and logs `[deep-link] failed to register formula:// handler: ...` if it fails. In environments where protocol registration is blocked/unreliable, prefer loopback redirects.
- **Loopback listener fails to start / port already in use:** pick a different port. The Rust command returns an error like `Failed to bind loopback OAuth redirect listener on 127.0.0.1:<port>: ...` (or `localhost:` / `[::1]:` depending on the host).
- **Provider uses port `0` (dynamic port selection):** not supported — the redirect URI must include an explicit, non-zero port.
- **Redirect is received but auth doesn’t complete:** ensure the redirect URI used in the auth request matches exactly (scheme/host/port/path). The frontend matcher is strict about `pathname` (e.g. `/callback` vs `/callback/`).
- **Using implicit flow (`#access_token` fragments):** loopback capture can only see query parameters; use auth-code + PKCE (code in the query string).

#### Tray + app menu + global shortcuts

- Tray menu and click behavior are implemented in `apps/desktop/src-tauri/src/tray.rs`.
  - Emits: `tray-new`, `tray-open`, `tray-quit`
  - “Check for Updates” runs an update check (`updater::spawn_update_check(..., UpdateCheckSource::Manual)`)
- App menu items are implemented in `apps/desktop/src-tauri/src/menu.rs` and forwarded as `menu-open` / `menu-save` / … events.
- In release builds, `main.rs` can run a startup update check, but it waits for the frontend to emit `updater-ui-ready` so update events aren’t dropped before listeners are installed.
- Global shortcuts are registered in `apps/desktop/src-tauri/src/shortcuts.rs`.
  - Accelerators: `CmdOrCtrl+Shift+O`, `CmdOrCtrl+Shift+P`
  - The plugin handler in `main.rs` emits: `shortcut-quick-open`, `shortcut-command-palette`

Note on quitting from the tray:

- The Rust host emits `tray-quit`, but it does **not** hard-exit immediately.
- The frontend handles `tray-quit` by running its quit flow (best-effort `Workbook_BeforeClose`, unsaved changes prompt) and finally invoking the `quit_app` command to exit the process.

---

## Frontend host wiring (`apps/desktop/src/main.ts`)

The desktop UI intentionally avoids a hard dependency on `@tauri-apps/api` and instead uses the injected runtime object:

- `globalThis.__TAURI__.core.invoke` for `#[tauri::command]` calls
- `globalThis.__TAURI__.event.listen` / `emit` for events
- `globalThis.__TAURI__.window.*` for hiding/showing/focusing the window
- `globalThis.__TAURI__.dialog.open/save` for file open/save prompts

Desktop-specific listeners are set up near the bottom of `apps/desktop/src/main.ts`:

- `oauth-redirect` → route deep-link redirects into the OAuth broker (buffers early redirects to avoid a rare PKCE race where the redirect arrives before `waitForRedirect` is registered); emits `oauth-redirect-ready` once the handler is installed (flushes queued redirects on the Rust side)
- `close-prep` → commit any in-progress edits (including split-view cell editors) + flush pending workbook sync + call `set_macro_ui_context` → emit `close-prep-done`
- `close-requested` → run `handleCloseRequest(...)` (unsaved changes prompt + hide vs quit) → emit `close-handled`
- `open-file` → queue workbook opens; then emits `open-file-ready` once the handler is installed (flushes any queued open-file requests on the Rust side; helper: `installOpenFileIpc` in `apps/desktop/src/tauri/openFileIpc.ts`)
- `file-dropped` → open the first dropped path
- `tray-open` / `tray-new` / `tray-quit` → open dialog/new workbook/quit flow
- menu events (e.g. `menu-open`, `menu-save`, `menu-quit`) → routed to the same “open/save/close” logic used by keyboard shortcuts and tray menu items
- Window-level keyboard shortcuts (desktop-only): `Cmd/Ctrl+N`, `Cmd/Ctrl+O`, `Cmd/Ctrl+S`, `Cmd/Ctrl+Shift+S`, `Cmd/Ctrl+W`, `Cmd/Ctrl+Q`
  - Implemented as **built-in keybindings** (`apps/desktop/src/commands/builtinKeybindings.ts`) routed through the **KeybindingService** (`apps/desktop/src/extensions/keybindingService.ts`).
  - These bindings execute `workbench.*` commands registered in the `CommandRegistry` (see `apps/desktop/src/commands/registerWorkbenchFileCommands.ts`), so they also surface in UI (Command Palette shortcut hints, etc.).
  - `when`-clauses are used for focus scoping (e.g. avoid firing while focus is in a text input / editor surface).
  - `SpreadsheetApp.onKeyDown(...)` checks `e.defaultPrevented` and treats the event as consumed to avoid double-handling when the desktop keybinding layer already claimed the shortcut.
- `shortcut-quick-open` / `shortcut-command-palette` → open dialog/palette
- updater events → handled by the updater UI (`apps/desktop/src/tauri/updaterUi.ts`)
  - `main.ts` emits `updater-ui-ready` once the updater listeners are installed (so the Rust host can safely start a startup update check in release builds).

Separately, startup metrics listeners are installed at the top of `main.ts` via:

- `installStartupTimingsListeners()` (listens for `startup:window-visible`, `startup:webview-loaded`, `startup:tti`, `startup:metrics`)
- `reportStartupWebviewLoaded()` (invokes the host command to emit the initial timing events)

Important implementation detail: invoke calls are serialized via `queueBackendOp(...)` / `pendingBackendSync` so that bulk edits (workbook sync) don’t race with open/save/close.

---

## Desktop IPC surface

### Commands (`#[tauri::command]` endpoints)

Most command handlers live in `apps/desktop/src-tauri/src/commands.rs`, with a few “shell” commands defined alongside the desktop host (e.g. in `apps/desktop/src-tauri/src/main.rs` and `apps/desktop/src-tauri/src/tray_status.rs`).

The command list is large; below are the “core” ones most contributors will interact with (not exhaustive):

- **Workbook lifecycle**
  - `open_workbook`, `new_workbook`, `save_workbook`, `mark_saved`, `add_sheet`, `rename_sheet`, `move_sheet`, `delete_sheet`
- **Cells / ranges / recalculation**
  - `get_cell`, `set_cell`, `get_range`, `set_range`, `recalculate`, `undo`, `redo`
  - Dependency inspection: `get_precedents`, `get_dependents`
  - Sheet bounds: `get_sheet_used_range`
- **Clipboard**
  - `clipboard_read`, `clipboard_write` (multi-format read/write: `text/plain`, `text/html`, `text/rtf`, `image/png`)
- **Workbook metadata (used by UI + Power Query + AI tooling)**
  - `get_workbook_theme_palette`, `list_defined_names`, `list_tables`
- **Pivot tables**
  - `create_pivot_table`, `refresh_pivot_table`, `list_pivot_tables`
- **Printing / export**
  - `get_sheet_print_settings`, `set_sheet_page_setup`, `set_sheet_print_area`, `export_sheet_range_pdf`
- **Local file access for Power Query sources (instead of Tauri FS plugin)**
  - `read_text_file`, `read_binary_file`, `read_binary_file_range`, `stat_file`, `list_dir`
- **Power Query secure storage + refresh state**
  - `power_query_cache_key_get_or_create`
  - `power_query_credential_get|set|delete|list`
  - `power_query_refresh_state_get|set`
  - `power_query_state_get|set`
- **OAuth (desktop redirect capture)**
  - `oauth_loopback_listen` (starts a temporary RFC 8252 loopback listener for redirect URIs using `http://127.0.0.1:<port>/...`, `http://localhost:<port>/...`, or `http://[::1]:<port>/...`; listener times out after 5 minutes)
- **SQL (connectors / queries)**
  - `sql_query`, `sql_get_schema`
- **Macros + scripting**
  - Macro inspection/security: `get_vba_project`, `list_macros`, `get_macro_security_status`, `set_macro_trust`
  - Execution/context: `set_macro_ui_context`, `run_macro`, `validate_vba_migration`
  - Python: `run_python_script`
  - VBA event hooks: `fire_workbook_open`, `fire_workbook_before_close`, `fire_worksheet_change`, `fire_selection_change`
- **Updates**
  - `check_for_updates` (triggers `updater::spawn_update_check(...)`; used by the command palette / manual update checks)
  - `install_downloaded_update` (installs the already-downloaded update bytes from the backend’s background download; used by the updater restart flow)
- **Lifecycle**
  - `quit_app` (hard-exits the process; used by the tray/menu quit flow)
  - `restart_app` (Tauri-managed restart/exit; intended for updater install flows so Tauri/plugins can shut down cleanly)
  - `--cross-origin-isolation-check` (special CLI mode; exits with 0/1 based on `crossOriginIsolated` + `SharedArrayBuffer`)
  - `--log-process-metrics` (prints a one-line host process snapshot: `[metrics] rss_mb=<n> pid=<pid>`)

Note: `quit_app` intentionally hard-exits (`std::process::exit(0)`) to avoid re-entering the hide-to-tray close handler.
For update-driven restarts prefer `restart_app` (graceful).
- **Tray integration**
  - `set_tray_status` (update tray icon + tooltip for simple statuses: `idle`, `syncing`, `error`)
- **Startup metrics**
  - `report_startup_webview_loaded`, `report_startup_tti`
- **Notifications**
  - `show_system_notification` (best-effort native notification via `tauri-plugin-notification`; used as a fallback by `apps/desktop/src/tauri/notifications.ts`, and restricted to the main window)
- **External URLs**
  - `open_external_url` (opens URLs in the OS via `tauri_plugin_shell`; allowlists `http:`, `https:`, `mailto:` and rejects everything else, including `javascript:`, `data:`, and `file:`; restricted to the main window + trusted app-local origins)

### Backend → frontend events

Events emitted by the Rust host (see `main.rs`, `menu.rs`, `tray.rs`, `updater.rs`):

- Window lifecycle:
  - `close-prep` (payload: token `string`)
  - `close-requested` (payload: `{ token: string, updates: CellUpdate[] }`)
  - `open-file` (payload: `string[]` paths)
  - `file-dropped` (payload: `string[]` paths)
- Deep links:
  - `oauth-redirect` (payload: `string` URL, e.g. `formula://oauth/callback?...`)
- Menu bar:
  - `menu-open`, `menu-new`, `menu-save`, `menu-save-as`, `menu-export-pdf`, `menu-close-window`, `menu-quit`
  - `menu-undo`, `menu-redo`, `menu-cut`, `menu-copy`, `menu-paste`, `menu-paste-special`, `menu-select-all`
  - `menu-zoom-in`, `menu-zoom-out`, `menu-zoom-reset`
  - `menu-about`, `menu-check-updates`, `menu-open-release-page`
- Tray:
  - `tray-new`, `tray-open`, `tray-quit`
- Shortcuts:
  - `shortcut-quick-open`, `shortcut-command-palette`
- Startup metrics:
  - `startup:window-visible` (payload: `number`)
  - `startup:webview-loaded` (payload: `number`)
  - `startup:tti` (payload: `number`)
  - `startup:metrics` (payload: `{ window_visible_ms?, webview_loaded_ms?, tti_ms? }`)
- Updates:
  - `update-check-started` (payload: `{ source }`)
  - `update-check-already-running` (payload: `{ source }`)
  - `update-not-available` (payload: `{ source }`)
  - `update-check-error` (payload: `{ source, message }`)
  - `update-available` (payload: `{ source, version, body }`)
  - `update-download-started` (payload: `{ source, version }`)
  - `update-download-progress` (payload: `{ source, version, chunkLength, downloaded, total?, percent? }`)
  - `update-downloaded` (payload: `{ source, version }`)
  - `update-download-error` (payload: `{ source, version, message }`)

Updater events are consumed by the desktop frontend in `apps/desktop/src/tauri/updaterUi.ts` (installed
from `apps/desktop/src/main.ts`).

Updater UX responsibilities:

- **Manual checks** (`source: "manual"`): show in-app feedback (toasts + focus the window), and show the
  update dialog when an update is available.
- **Startup checks** (`source: "startup"`): show a **system notification only** when an update is available
  (no in-app dialog). The backend also starts a best-effort background download; once it completes
  (`update-downloaded`), the frontend shows an in-app “Update ready to install” toast so the user can
  restart/apply when convenient.
  - If the user triggers "Check for Updates" while a startup check is already in-flight, the backend may
    later emit a completion event with `source: "startup"`. The frontend treats that result as manual UX
    so the user still sees the expected dialog/toasts.

The update-available dialog includes an **"Open release page"** action that opens the GitHub Releases page
for manual downgrade/rollback. If an update download or install fails, this action is relabeled/promoted
to **"Download manually"** and the dialog’s status text includes manual download/downgrade instructions.

Related frontend → backend events used as acknowledgements / readiness signals:

- `close-prep-done` (token)
- `close-handled` (token)
- `open-file-ready` (signals that the frontend’s `open-file` listener is installed; causes the Rust host to flush queued open requests)
- `oauth-redirect-ready` (signals that the frontend’s `oauth-redirect` listener is installed; causes the Rust host to flush queued OAuth/deep-link redirects)
- `updater-ui-ready` (signals the updater UI listeners are installed; triggers the startup update check)
- `coi-check-result` (used only by the cross-origin isolation smoke check mode; see `pnpm -C apps/desktop check:coi`)

Security note: these event names are **explicitly allowlisted** in
`apps/desktop/src-tauri/capabilities/main.json`. If you add a new desktop event, you must update
that allowlist (and the guardrail test `apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts`)
or the event will fail with a permissions error in hardened desktop builds.

Manual verification: in a Tauri desktop build, try calling `__TAURI__.event.listen(...)` or
`__TAURI__.event.emit(...)` with an event name that is **not** in the allowlist; the call should
be rejected with a permissions error.

---

## Clipboard

Clipboard read/write is implemented as a **platform provider** so the same copy/paste code paths
work on both desktop (Tauri) and web.

Frontend entry point:

- `apps/desktop/src/clipboard/platform/provider.js`

Rust implementation:

- Tauri commands: `apps/desktop/src-tauri/src/clipboard/mod.rs` (`clipboard_read`, `clipboard_write`)
- Legacy clipboard commands (fallback path / main-thread bridging on macOS): `apps/desktop/src-tauri/src/commands.rs` (`read_clipboard`, `write_clipboard`)
- Platform backends: `apps/desktop/src-tauri/src/clipboard/platform/*` (delegates into OS-specific modules like `clipboard/macos.rs`)
- Windows helpers:
  - CF_HTML encode/decode for the `"HTML Format"` clipboard format: `apps/desktop/src-tauri/src/clipboard/cf_html.rs`
  - PNG ↔ DIBV5 conversion for image clipboard interop: `apps/desktop/src-tauri/src/clipboard/windows_dib.rs`
- Linux clipboard fallback heuristics (X11 `PRIMARY` vs `CLIPBOARD`): `apps/desktop/src-tauri/src/clipboard_fallback.rs`

Provider selection:

- `createClipboardProvider()` chooses **Tauri vs web** by checking for `globalThis.__TAURI__`.

End-to-end flow (grid copy/paste):

1. UI handlers in `apps/desktop/src/app/spreadsheetApp.ts` intercept `Cmd/Ctrl+C`, `X`, `V`.
2. Copy/cut builds a `CellGrid` via `getCellGridFromRange()` and serializes it via `serializeCellGridToClipboardPayload()` (`apps/desktop/src/clipboard/clipboard.js`), producing `{ text, html, rtf }`.
   - For formula cells, the copy/cut path ensures `text/plain` contains the **displayed value** (computed formula result), while `text/html` can still preserve formulas (via `data-formula`) for spreadsheet-to-spreadsheet pastes.
3. The platform provider (`apps/desktop/src/clipboard/platform/provider.js`) writes/reads the system clipboard.
4. Paste parses clipboard payloads in priority order: `text/html` → `text/plain` → `text/rtf` (plain-text extracted from RTF).

Desktop vs web behavior:

- **Desktop (Tauri)**: prefers custom Rust commands `clipboard_read` / `clipboard_write` for
  **rich, multi-format** clipboard access via `globalThis.__TAURI__.core.invoke(...)`.
  - If `clipboard_read` is missing or errors (older builds / unsupported platforms / threading constraints),
    the provider also tries the legacy command name `read_clipboard` as a best-effort merge (never
    clobbering WebView values).
  - If `clipboard_write` is missing (older builds) or errors, the provider also tries the legacy
    command name `write_clipboard` before falling back to plain-text clipboard APIs.
  - If native commands are unavailable/unimplemented, it falls back to the Web Clipboard API
    (`navigator.clipboard`) and finally to `globalThis.__TAURI__.clipboard.readText` / `writeText`
    (plain text).
- **Web**: uses the browser Clipboard API only (permission + user-gesture gated).

Supported MIME types (read + write, best-effort):

- `text/plain`
- `text/html`
- `text/rtf`
- `image/png`

Size limits (defense-in-depth):

- Native clipboard reads/writes may encounter extremely large payloads (notably screenshots). To keep paste responsive and avoid huge IPC/base64 transfers, the desktop app applies best-effort size caps:
  - `image/png`: **5 MiB** (raw PNG bytes; base64 over IPC is larger)
  - `text/plain`, `text/html`, `text/rtf`: **2 MiB** (UTF-8 bytes)
- Oversized formats are **omitted** on read (no error).
- Clipboard writes validate payload sizes and may omit/cap rich formats (or fail validation) depending on the code path.

JS provider contract (normalized API):

- Read: `provider.read()` → `{ text?: string, html?: string, rtf?: string, imagePng?: Uint8Array, pngBase64?: string }`
- Write: `provider.write({ text, html?, rtf?, imagePng?, pngBase64? })` → `void`

Notes:

- `imagePng` (raw bytes) is the primary JS-facing image API.
- `pngBase64` is a legacy/internal escape hatch. The provider will generally decode any native base64 payload into `imagePng` before returning it to callers, and only preserves `pngBase64` when decoding fails.

Tauri wire contract (internal, used only for `__TAURI__.core.invoke`):

- Read: `invoke("clipboard_read")` → `{ text?: string, html?: string, rtf?: string, image_png_base64?: string }`
- Write: `invoke("clipboard_write", { payload: { text?, html?, rtf?, image_png_base64? } })` → `void`
- Legacy read fallback: `invoke("read_clipboard")` → `{ text?: string, html?: string, rtf?: string, image_png_base64?: string }`
- Legacy write fallback: `invoke("write_clipboard", { text, html?, rtf?, image_png_base64? })` → `void`

Provider return shape:

- `createClipboardProvider().read()` returns a merged `ClipboardContent` and normalizes images to `imagePng: Uint8Array` when possible.
- `pngBase64` may be present only as a legacy/internal fallback when decoding into bytes fails; callers should not rely on it.

Image wire format:

- JS-facing APIs use **raw PNG bytes** (`imagePng: Uint8Array`).
- Over Tauri IPC, PNG bytes are transported as a base64 string (**raw base64**, no `data:image/png;base64,` prefix).
  - The canonical key for this repo is `image_png_base64`, but the frontend provider also tolerates legacy aliases (`pngBase64`, `png_base64`) from older builds / bridges.
- The platform provider (`apps/desktop/src/clipboard/platform/provider.js`) is responsible for converting base64 ↔ bytes when crossing the IPC boundary.

Known platform limitations:

- **Web Clipboard API permission gating**: `navigator.clipboard.read/write` is user-gesture gated and may be denied by the OS/WebView permission model.
- **HTML/RTF availability varies** by WebView and OS: some platforms allow reading `text/html`/`text/rtf` but deny writing them (or vice versa).
- **Image clipboard support varies**: `image/png` via `ClipboardItem` is not consistently supported across all embedded WebViews.
- **Linux selection semantics**: when running under X11, if `CLIPBOARD` has no usable content, the native GTK clipboard backend may fall back to `PRIMARY` selection (middle-click paste). This fallback is skipped on Wayland by default, but can be overridden via `FORMULA_CLIPBOARD_PRIMARY_SELECTION=0/false/no` (disable) or `=1/true/yes` (force-enable). The same setting also gates whether Formula populates `PRIMARY` on copy/cut.
- The native clipboard commands are implemented per-OS; when they are missing or unimplemented, the provider falls back to Web Clipboard (and then plain text).

Security boundaries:

- **IPC origin hardening (defense-in-depth):** clipboard commands (`clipboard_read`, `clipboard_write`, and legacy
  `read_clipboard` / `write_clipboard`) enforce **main-window + trusted app origin** checks in Rust via
  `apps/desktop/src-tauri/src/ipc_origin.rs` (trusted: `localhost` / `*.localhost` / `127.0.0.1` / `::1`, best-effort
  `file://`; denied: remote hosts, `data:`).
- **DLP enforcement happens before writing**: grid copy/cut paths perform DLP checks before touching the system clipboard:
  - `SpreadsheetApp.copySelectionToClipboard()` / `cutSelectionToClipboard()`
  - → `enforceClipboardCopy` (`apps/desktop/src/dlp/enforceClipboardCopy.js`)
- **Extensions are DLP-enforced too**: `formula.clipboard.writeText(...)` is mediated by the desktop extension host adapter and
  enforces clipboard-copy DLP before writing to the system clipboard. Enforcement considers both the current UI selection
  (active-cell fallback) and any spreadsheet ranges the extension has read/received (taint tracking).
- **Extension sandboxing**: extension panels run in sandboxed iframes; do not expose Tauri IPC (`invoke`) or native clipboard APIs directly to untrusted iframe content. Clipboard operations must be mediated by the trusted host UI layer.

### Debugging / troubleshooting

Clipboard interop bugs are often **format- and OS-dependent**. The desktop app provides opt-in debug logs on both the Rust (native) and JS (provider) sides.

**Rust (native clipboard backend)**: set `FORMULA_DEBUG_CLIPBOARD=1` **before launching** the desktop app (it’s read at process startup).

Dev (`tauri dev`) example:

```bash
cd apps/desktop
FORMULA_DEBUG_CLIPBOARD=1 bash ../../scripts/cargo_agent.sh tauri dev
```

In packaged/release builds, launch the installed desktop binary with `FORMULA_DEBUG_CLIPBOARD=1` set in the environment to get the same `[clipboard] …` log lines.

**JS provider (frontend)**:

- For a running session, enable logs in DevTools:
  ```js
  globalThis.FORMULA_DEBUG_CLIPBOARD = true;
  // or:
  globalThis.__FORMULA_DEBUG_CLIPBOARD__ = true;
  ```
- Optional build-time flag (Vite): set `VITE_FORMULA_DEBUG_CLIPBOARD=1` when building/running the desktop frontend so `import.meta.env.VITE_FORMULA_DEBUG_CLIPBOARD` is truthy.

These debug logs include **format names/sources and byte counts only** (no clipboard contents). Still, avoid collecting or sharing clipboard diagnostics that could include sensitive user data.

### Manual QA matrix (recommended)

| Platform | Copy from Formula | Paste target | What to verify |
|----------|-------------------|--------------|----------------|
| Windows | A range with formatting / an HTML table | Excel, Word | Table structure preserved (HTML), values align, formatting is reasonable |
| Windows | A copied chart/screenshot (PNG) | PowerPoint, Slack | Image pastes as an image (not a file path / empty paste) |
| macOS | A range with formatting / an HTML table | Notes, Pages | Table pastes with expected structure and styling |
| Linux | A range with formatting / an HTML table | LibreOffice + browser | HTML/table paste where supported; plain text fallback otherwise |

---

## Cargo feature gating (`desktop`)

The Tauri **binary** is feature-gated so that backend unit tests can run without system WebView
dependencies (notably GTK/WebKit on Linux).

Where it’s defined:

- `apps/desktop/src-tauri/Cargo.toml`
  - The desktop binary (`[[bin]]`) has `required-features = ["desktop"]`.
  - The `desktop` feature enables the optional deps: `tauri`, `tauri-build`, and the desktop-only Tauri plugins
    (currently `tauri-plugin-global-shortcut`, `tauri-plugin-single-instance`, `tauri-plugin-notification`, `tauri-plugin-dialog`,
    `tauri-plugin-clipboard-manager`, `tauri-plugin-shell`, `tauri-plugin-updater`, `tauri-plugin-deep-link`), plus a few desktop-only helpers (e.g.
    `http-range`, `percent-encoding`), plus Linux-only GTK deps for the clipboard backend, and enables the `parquet` feature.
- `apps/desktop/src-tauri/tauri.conf.json`
  - `build.features: ["desktop"]` ensures the desktop binary is compiled with the correct feature set when running the Tauri CLI
    (in agent environments: `cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri dev|build`).

Practical effect:

- Backend/unit tests can run in CI without installing WebView toolchains:
  - `bash scripts/cargo_agent.sh test -p desktop`
- Validating the full desktop build locally requires the platform WebView dependencies:
  - `bash scripts/cargo_agent.sh check -p desktop --features desktop`

Note: most `#[tauri::command]` functions in `apps/desktop/src-tauri/src/commands.rs` are also `#[cfg(feature = "desktop")]`, so the
backend library can still compile (and be tested) without linking Tauri or system WebView components.

---

## Tauri v2 Capabilities & Permissions

In Tauri v2, permissioning is driven by **capabilities** rather than the Tauri v1 “allowlist”.

Source of truth in this repo:

- `apps/desktop/src-tauri/capabilities/main.json`

Capabilities are scoped per window by the capability file’s `"windows": [...]` list (window labels from
`apps/desktop/src-tauri/tauri.conf.json`).

Note: some Tauri toolchains support window-level opt-in via `app.windows[].capabilities`, but the current toolchain used in
this repo rejects that field. Keep capability scoping in the capability file itself.

### What `main.json` does

`apps/desktop/src-tauri/capabilities/main.json` is intentionally an explicit allowlist for what the webview is allowed to do.

It gates:
- **`allow-invoke`** (application permission): allows the frontend to invoke Formula's app-defined Rust `#[tauri::command]` functions via
  `__TAURI__.core.invoke(...)`.
  - The command allowlist lives in `apps/desktop/src-tauri/permissions/allow-invoke.json`.
  - This allowlist should match the backend’s exposed command surface (`generate_handler![...]`), guardrailed by
    `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs`.
  - Even with allowlisting, commands must validate scope/authorization in Rust (trusted-origin + window-label checks,
    argument validation, filesystem/network scope checks, etc).
- **`core:allow-invoke`** (scoped core permission): a second, capability-local command allowlist.
  - In this repo it is granted in `apps/desktop/src-tauri/capabilities/main.json` using the **object form** with an explicit list:
    `{ "identifier": "core:allow-invoke", "allow": [{ "command": "..." }] }`.
  - We never grant the string form `"core:allow-invoke"` (it behaves like an unscoped/default allowlist).
  - Keep it explicit (no wildcards) and in sync with actual frontend `invoke("...")` usage (guardrailed by
    `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`).
- **`core:event:allow-listen` / `core:event:allow-emit`**: which event names the frontend can `listen(...)` for or `emit(...)`.
- **`core:event:allow-unlisten`**: allows the frontend to unregister event listeners it previously installed (so we don’t leak
  listeners for one-shot flows like close/open/OAuth readiness signals).
- Additional core/plugin permissions for using JS plugin APIs (dialog/window/clipboard/updater), for example:
  - `dialog:allow-open`, `dialog:allow-save`, `dialog:allow-confirm`, `dialog:allow-message`
  - `core:window:allow-hide`, `core:window:allow-show`, `core:window:allow-set-focus`, `core:window:allow-close`
  - `clipboard-manager:allow-read-text`, `clipboard-manager:allow-write-text`
  - `updater:allow-check`, `updater:allow-download`, `updater:allow-install`

Custom Rust commands (everything behind `#[tauri::command]`, invoked via `__TAURI__.core.invoke(...)`) are allowlisted by
`allow-invoke`, but must still be hardened in Rust (window label + trusted-origin checks, argument validation,
filesystem/network scope checks, etc.).

Note: the clipboard plugin permissions above only cover the legacy **plain-text** clipboard helpers
(`globalThis.__TAURI__.clipboard.readText` / `writeText`). Rich clipboard formats (HTML/RTF/PNG) are handled via
custom Rust IPC commands (`clipboard_read` / `clipboard_write`) and must validate/scoped in Rust.

External URL opening is also routed through a custom Rust IPC command:

- `open_external_url`
  - Enforces a strict scheme allowlist in Rust (`http:`, `https:`, `mailto:`).
  - Rejects dangerous/unsupported schemes (`javascript:`, `data:`, `file:`, and anything else).
  - Only callable from the **main** window and from trusted app-local origins (to prevent navigated-to remote content from using IPC to open links).

Note: this capability intentionally does **not** grant `shell:allow-open` (the JS shell plugin API). Prefer using the `open_external_url`
command (via `apps/desktop/src/tauri/shellOpen.ts`) so link handling remains consistent and scheme allowlisting lives at a
single trusted boundary.

High-level contents (see the file for the exhaustive list):

- We avoid `core:default` (broad, unscoped access to core plugins like event/window) to keep the permission surface minimal/explicit.
- We keep custom Rust IPC calls explicit via:
  - `allow-invoke` (application permission defined in `apps/desktop/src-tauri/permissions/allow-invoke.json`, kept in sync with `generate_handler![...]`)
  - `core:allow-invoke` (core permission object in `apps/desktop/src-tauri/capabilities/main.json`, kept in sync with real `invoke("...")` usage)
- We scope `core:event:allow-listen` / `core:event:allow-emit` to explicit event-name allowlists (no wildcards).
- `core:event:allow-listen` includes:
  - close flow: `close-prep`, `close-requested`
  - open flow: `open-file`, `file-dropped`
  - deep links / OAuth: `oauth-redirect`
  - native menu bar events (e.g. `menu-open`, `menu-save`, `menu-check-updates`)
  - tray + shortcuts (e.g. `tray-open`, `shortcut-command-palette`)
  - startup timing instrumentation (e.g. `startup:webview-loaded`, `startup:tti`)
  - updater events (e.g. `update-check-started`, `update-available`)
- `core:event:allow-emit` includes acknowledgements and check-mode signals:
  - `open-file-ready`, `oauth-redirect-ready`
  - `close-prep-done`, `close-handled`
  - `updater-ui-ready`
  - `coi-check-result` (used by `pnpm -C apps/desktop check:coi`)
- `core:event:allow-unlisten` is granted so the frontend can clean up its own temporary listeners.
- Plugin permissions include dialog/window/clipboard APIs plus updater permissions (`updater:allow-check`, `updater:allow-download`, `updater:allow-install`, required for the updater UI).
  - Window API permissions are `core:window:allow-*`.
  - Plain-text clipboard permissions are `clipboard-manager:allow-*`.
- Custom Rust commands are allowlisted by `allow-invoke`, but must still keep input validation and scope checks in Rust.

We intentionally keep capabilities narrow and rely on explicit Rust commands + higher-level app permission gates (macro
trust, DLP, extension permissions) for privileged operations.

Guardrail tests (to prevent accidental “allow everything” capability drift):

- `apps/desktop/src/tauri/__tests__/tauriSecurityConfig.vitest.ts` — asserts the hardened CSP/headers (COOP/COEP, no framing, restricted network) and capability scoping:
  - `capabilities/main.json` includes `"windows": ["main"]`
  - if `app.windows[].capabilities` is present, the `main` window includes `"capabilities": ["main"]` (and no other window has `main`)
  - otherwise, no window should specify `capabilities` (toolchain compatibility)
- `apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts` — asserts the `core:event:allow-listen` / `core:event:allow-emit`
  allowlists match the desktop shell’s real event usage (and contain no wildcards).
- `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts` — asserts required plugin permissions stay explicit/minimal (dialogs, window ops, clipboard plain text, updater, etc), we don’t grant dangerous extras (e.g. `shell:allow-open`, notification permissions), and that `allow-invoke.json` stays explicit and in sync with frontend invoke usage (no allow-all).
- `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs` — asserts the `allow-invoke` permission allowlist stays in sync with the `generate_handler![...]` list in `apps/desktop/src-tauri/src/main.rs`.
- `apps/desktop/src/tauri/__tests__/openFileIpcWiring.vitest.ts` — asserts the open-file IPC handshake (`open-file-ready`) is still wired in `main.ts` (prevents cold-start open drops).
- `apps/desktop/src/tauri/__tests__/updaterMainListeners.vitest.ts` — asserts updater UX listeners remain consolidated in `tauri/updaterUi.ts` and the `updater-ui-ready` handshake stays intact.

### Validating permissions against the Tauri toolchain

When upgrading Tauri or plugins, the set of valid permission identifiers can change. You can validate the capability
files against the **actual** toolchain installed in your environment by regenerating the schemas and listing the
available permissions:

```bash
# Generates `apps/desktop/src-tauri/gen/schemas/desktop-schema.json` (ignored by git).
bash scripts/cargo_agent.sh check -p desktop --features desktop --lib

# Lists all available `${plugin}:${permission}` identifiers.
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri permission ls
```

Note: on Tauri v2.9, core permissions use the `core:` prefix (e.g. `core:event:allow-listen`, `core:window:allow-hide`).

### Practical workflow

- If you add a new event name used by `listen(...)` or `emit(...)`, update the `core:event:allow-listen` / `core:event:allow-emit`
  allowlists (in `apps/desktop/src-tauri/capabilities/main.json`).
- If the frontend starts using a new Tauri core/plugin API (dialog/window/clipboard/updater), add the corresponding `*:allow-*`
  permission string(s).
  - Window permissions are currently `core:window:allow-*`.
  - Plain-text clipboard permissions are currently `clipboard-manager:allow-*`.
- For custom Rust `#[tauri::command]` functions invoked via `__TAURI__.core.invoke(...)`:
  - register them in `apps/desktop/src-tauri/src/main.rs` (`generate_handler![...]`)
  - add them to `apps/desktop/src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission `commands.allow`)
  - add them to the `core:allow-invoke` allowlist in `apps/desktop/src-tauri/capabilities/main.json` (`allow: [{ "command": "..." }]`)
  - keep input validation and scope checks in Rust (trusted-origin + window-label checks, etc)

Guardrails (CI/tests):

- `apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts` enforces that the event allowlists are explicit (no allow-all) and match the events used by the desktop code.
- `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts` asserts the `allow-invoke.json` command allowlist stays explicit and in sync with frontend invoke usage. It also keeps the plugin permission surface minimal/explicit (including split updater `allow-check` / `allow-download` / `allow-install`).

For background on capability syntax/semantics, see the upstream Tauri v2 docs:

- https://tauri.app/v2/guides/security/capabilities/

---

## Filesystem access + scope enforcement

### Why we use custom commands (Power Query)

The desktop webview needs local filesystem access for Power Query sources (CSV/JSON/Parquet, folder listings).

This repo does **not** require the official Tauri FS plugin to be enabled: instead, it uses custom Rust commands in:

- `apps/desktop/src-tauri/src/commands.rs`

Notable commands:

- `read_text_file`
- `read_binary_file`
- `read_binary_file_range`
- `stat_file`
- `list_dir`

### Scope enforcement

These commands enforce a filesystem scope equivalent to the platform allowlist:

- `$HOME/**`
- `$DOCUMENT/**`
- `$DOWNLOADS/**` (if the OS/user has a Downloads dir configured and it exists/canonicalizes successfully; notably on Linux this may be outside `$HOME` via XDG user dirs)

Implementation notes:

- The scope helper lives in `apps/desktop/src-tauri/src/fs_scope.rs`.
- Commands additionally enforce **main-window + trusted app origin** checks via `apps/desktop/src-tauri/src/ipc_origin.rs`
  (defense-in-depth so remote/untrusted navigations can't invoke privileged filesystem reads).
- Requested paths are **canonicalized** before checking scope.
- Canonicalization normalizes `..` traversal and prevents symlink escapes (e.g. a symlink inside `$HOME` pointing to `/etc/passwd` is rejected).
- `list_dir` validates the root directory and validates individual entries before returning metadata.

---

## Release, signing, updater keys

The updater config (`plugins.updater.*`) is in `apps/desktop/src-tauri/tauri.conf.json`.

For the actual release workflow, signing, and updater key management, see:

- `docs/release.md`
