# Formula Desktop (Vite/Webview)

## Desktop perf (startup, memory, size)

From the repo root, you can run the desktop perf helpers (builds the Tauri binary + frontend as needed):

```bash
pnpm perf:desktop-startup
pnpm perf:desktop-memory
pnpm perf:desktop-size
```

These commands run the desktop app with an isolated, repo-local HOME directory under `target/perf-home` by default
(override via `FORMULA_PERF_HOME`, preserve via `FORMULA_PERF_PRESERVE_HOME=1`).

More details (including metric definitions and CI gating env vars):

- `docs/11-desktop-shell.md`
- `docs/16-performance-targets.md`

## Bundle analysis (renderer build)

To inspect which chunks/dependencies dominate the desktop renderer bundle, run:

```bash
pnpm -C apps/desktop build:analyze
```

This enables `rollup-plugin-visualizer` via `VITE_BUNDLE_ANALYZE=1` and writes the report(s) to:

- `apps/desktop/dist/bundle-stats.html` (interactive treemap)
- `apps/desktop/dist/bundle-stats.json` (raw data)

Note: `pnpm -C apps/desktop build:analyze` sets `VITE_BUNDLE_ANALYZE=1` automatically.

Optional: for more accurate *per-module* attribution, you can also enable sourcemap-based analysis:

```bash
pnpm -C apps/desktop build:analyze:sourcemap
```

This additionally generates `apps/desktop/dist/bundle-stats-sourcemap.html` (and enables `build.sourcemap` for that run).

To get a quick CLI summary of the largest chunks and dependency groups (from `bundle-stats.json`):

```bash
pnpm -C apps/desktop report:bundle-stats
```

To focus on the **startup path** (only the JS referenced by `dist/index.html` via `<script>` + `<link rel="modulepreload">`):

```bash
pnpm -C apps/desktop report:bundle-stats -- --startup
```

Normal builds (`pnpm -C apps/desktop build`) are unchanged unless `VITE_BUNDLE_ANALYZE=1` is set.

## JS bundle size budgets (CI guard)

CI runs a lightweight JS bundle size check after the desktop Vite build to prevent accidental
dependency additions that bloat the desktop initial load.

Locally:

```bash
pnpm -C apps/desktop build
pnpm -C apps/desktop check:bundle-size
```

The check prints a markdown summary and (optionally) enforces budgets via env vars (KiB = 1024 bytes):

- `FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB` – total JS in `apps/desktop/dist/**/*.js`
- `FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB` – entry JS referenced by `apps/desktop/dist/index.html` `<script>` tags
- `FORMULA_DESKTOP_JS_ASSETS_BUDGET_KB` (optional) – total Vite JS in `apps/desktop/dist/assets/**/*.js`

Optional:

- `FORMULA_DESKTOP_BUNDLE_SIZE_WARN_ONLY=1` – report budget violations but exit 0
- `FORMULA_DESKTOP_BUNDLE_SIZE_SKIP_GZIP=1` – skip gzip computation (faster)

## Frontend asset download size (compressed JS/CSS/WASM)

To measure the **network download size** of the frontend (Vite `dist/assets`) using Brotli or gzip:

```bash
pnpm -C apps/desktop build
node scripts/frontend_asset_size_report.mjs --dist apps/desktop/dist
```

Optional budget enforcement (MB = 1,000,000 bytes):

- `FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB=10` (default: 10MB total)
- `FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=brotli|gzip` (default: brotli)
- `FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1` to fail when the total exceeds the limit

## Dist asset breakdown (static assets)

The desktop build can include large static assets (Pyodide being the biggest when bundled), which can dominate the
packaged app footprint even when JS bundle sizes are stable. By default, Pyodide is downloaded on-demand at runtime
and is **not** included in `dist/` unless bundling is explicitly enabled.

To see which files contribute most to `apps/desktop/dist/`, run:

```bash
pnpm build:desktop
node scripts/desktop_dist_asset_report.mjs
```

Alternatively, from `apps/desktop/`:

```bash
pnpm build
pnpm report:dist-assets
```

This prints a Markdown report of:

- total `dist/` size
- the **largest files** (top 25 by default; use `--top N`)
- **grouped totals** by directory prefix (top-level by default; use `--group-depth N`)
- **file type totals** (by extension; disable with `--no-types`)

Optional budget enforcement (MB = 1,000,000 bytes):

- `FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB` – fail if total `dist/` size exceeds this value
- `FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB` – fail if any single file exceeds this value

Optional: write a machine-readable JSON report (still prints Markdown to stdout):

```bash
node scripts/desktop_dist_asset_report.mjs --json-out desktop-dist-assets.json
```

## Rust binary size analysis (desktop shell)

To inspect which Rust **crates/symbols** dominate the `formula-desktop` release binary (useful for bundle-size
optimization work), run:

```bash
# Recommended analysis tool
cargo install cargo-bloat --locked

# Build (release) + report
python3 scripts/desktop_binary_size_report.py

# If you've already built the release binary, you can skip the build step:
python3 scripts/desktop_binary_size_report.py --no-build
```

The report is emitted as Markdown to stdout, and is also published to the GitHub Actions step summary in the
`Performance` workflow.

Tip: `pnpm perf:desktop-size` also includes this Rust binary breakdown alongside the existing dist/bundle size summaries.

### Optional: enforce a size budget

The report is informational by default. To turn it into a regression gate (locally or in CI), set:

```bash
export FORMULA_DESKTOP_BINARY_SIZE_LIMIT_MB=XX
export FORMULA_ENFORCE_DESKTOP_BINARY_SIZE=1
```

Or pass `--limit-mb` / `--enforce` to `scripts/desktop_binary_size_report.py`.

## Workbook load limits (snapshot loading)

When opening a workbook in the desktop app, the renderer fetches a **cell snapshot** from the backend to populate the UI.
To avoid excessive memory/time costs for very large workbooks, snapshot loading is capped by default:

- **Rows:** first `10,000`
- **Columns:** first `200`

If the workbook’s used range exceeds these limits, the app will show a **warning toast** indicating that only a prefix of the
workbook was loaded.

### Overriding the limits

Preferred (runtime):

- URL query params: `?loadMaxRows=<n>&loadMaxCols=<n>` (e.g. `?loadMaxRows=50000&loadMaxCols=1000`)
  - Backwards-compat: `?maxRows=<n>&maxCols=<n>` is still accepted.
- Optional: control snapshot fetch chunk size with `?loadChunkRows=<n>` (back-compat: `?chunkRows=<n>`).

Persistent (runtime):

- localStorage keys:
  - `formula.desktop.workbookLoadMaxRows`
  - `formula.desktop.workbookLoadMaxCols`
  - `formula.desktop.workbookLoadChunkRows` (optional)

Fallback (build/dev-time):

- Environment variables:
  - Vite/WebView runtime (recommended): `VITE_DESKTOP_LOAD_MAX_ROWS` / `VITE_DESKTOP_LOAD_MAX_COLS` (optional: `VITE_DESKTOP_LOAD_CHUNK_ROWS`)
  - Node/tooling/tests: `DESKTOP_LOAD_MAX_ROWS` / `DESKTOP_LOAD_MAX_COLS` (optional: `DESKTOP_LOAD_CHUNK_ROWS`)

Invalid / non-positive values are ignored and fall back to the defaults above.

## Collaboration (real-time sync dev)

The desktop app can run in a simple real-time collaboration mode backed by the local sync server (`services/sync-server`)
using Yjs + `@formula/collab-session`.

### Start the sync server

```bash
pnpm dev:sync
```

The dev server defaults to `ws://127.0.0.1:1234` and accepts the default dev token `dev-token`.

### Start the desktop dev server

```bash
pnpm -C apps/desktop dev
```

### Open two clients

Each desktop client gets a **stable collaboration identity** (`id`/`name`/`color`) persisted in localStorage
(`formula:collab:user`). For local testing, you can override the identity per-window via URL query params
(`collabUserId`, `collabUserName`, `collabUserColor`).

Open two browser windows pointing at the same `docId` with different user identities:

Window 1:

```
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=demo&token=dev-token&collabUserId=u1&collabUserName=Alice&collabUserColor=%234c8bf5
```

Window 2:

```
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=demo&token=dev-token&collabUserId=u2&collabUserName=Bob&collabUserColor=%23f97316
```

### Dev: exercise end-to-end cell encryption

To test encrypted cell payloads (`enc`) end-to-end, open one client with `collabEncrypt=1` (and optionally `collabEncryptRange=...`) and another client without it.

Example (client A has keys, client B does not):

```text
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=demo&token=dev-token&collabUserId=u1&collabUserName=Alice&collabUserColor=%234c8bf5&collabEncrypt=1&collabEncryptRange=Sheet1!A1:C10
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=demo&token=dev-token&collabUserId=u2&collabUserName=Bob&collabUserColor=%23f97316
```

Local persistence is enabled by default (IndexedDB). To disable it for debugging, add `&collabPersistence=0`.

Edits and comments should sync in real-time.

The status bar shows a collaboration indicator (`Collab: …`) with:

- the current `docId`
- websocket connectivity (`Connecting…` / `Connected` / `Disconnected` / `Offline`)
- sync state (`Syncing…` / `Synced`)

To exercise the conflict UI, edit the same cell concurrently in two windows (e.g. different formulas in the same cell).

## Pyodide / Python scripting

The Pyodide-based Python runtime (`@formula/python-runtime`) supports two internal backends:

- **Worker backend (preferred)**: runs Pyodide in a Worker and uses `SharedArrayBuffer + Atomics` to keep spreadsheet RPC synchronous.
  This requires a **cross-origin isolated** context (COOP/COEP).
- **Main-thread backend (fallback)**: runs Pyodide on the main thread and calls the spreadsheet bridge synchronously.
  This works in non-COOP/COEP contexts but may freeze the UI while Python runs.

In `mode: "auto"` (default), the runtime prefers the Worker backend when possible and falls back to the main thread otherwise.

This app’s Vite dev/preview servers are configured to enable that:

- `Cross-Origin-Opener-Policy: same-origin`
- `Cross-Origin-Embedder-Policy: require-corp`

For packaged (production) Tauri builds, the same headers are configured via
`app.security.headers` in `apps/desktop/src-tauri/tauri.conf.json` (see `docs/11-desktop-shell.md`).

Packaged desktop builds download Pyodide assets **on-demand** into an app-data
cache and serve them to the WebView via the `pyodide://` protocol (COOP/COEP
friendly). This keeps installers small while preserving Python functionality.

When running the frontend outside the desktop shell (e.g. in a browser via the
Vite dev/preview servers), Pyodide defaults to loading from the official CDN.
The CDN provides COEP/CORP-friendly headers so this still works in a
cross-origin isolated context.

In **packaged/production desktop builds**, Pyodide assets are **not embedded** in `dist/` by
default to keep installer size down. The first time the user runs a Pyodide-backed feature, the
desktop backend downloads + verifies the pinned Pyodide distribution into the app cache directory
and serves it via the `pyodide://` protocol.

Security note: in packaged desktop builds, `__pyodideIndexURL` overrides are ignored unless they
point at a local origin (`pyodide://...` or `/pyodide/...`). This avoids loading an arbitrary Python
runtime from the network.

If you need to self-host/bundle Pyodide under `/pyodide/v0.25.1/full/` (for
offline development/CI/preview), set `FORMULA_BUNDLE_PYODIDE_ASSETS=1` when
running `pnpm -C apps/desktop dev` or `pnpm -C apps/desktop build` (this runs
`scripts/ensure-pyodide-assets.mjs`, populating `apps/desktop/public/pyodide/`
which Vite then serves/copies into `dist/`).

## Content Security Policy (Tauri)

The desktop app ships with a strict CSP in `apps/desktop/src-tauri/tauri.conf.json`.

In packaged Tauri builds, `connect-src` allows outbound network for:

- HTTPS (`https:`)
- WebSockets (`ws:`/`wss:`) — required for collaboration (Yjs via `y-websocket`)
- Local `blob:`/`data:` URLs used by the extension system

The extensions + marketplace runtime prefer using Rust-backed Tauri IPC commands for outbound HTTP(S):

- `network_fetch` — used by the browser extension host for `formula.network.fetch(...)`
- `marketplace_*` — used by the marketplace client

This avoids relying on browser CORS behavior and keeps network policy centralized (see `docs/11-desktop-shell.md` → “Network strategy”).

Note: Rust IPC network requests are not governed by the WebView CSP. `network_fetch` / `marketplace_*` currently allow
`http:` URLs as well as `https:` (useful for local dev servers), even though `connect-src` still does not include `http:`.

To fetch the assets without starting Vite:

```bash
pnpm -C apps/desktop pyodide:setup
```

For an end-to-end smoke test, open:

- `/python-runtime-test.html` (PyodideRuntime + formula bridge)
- `/scripting-test.html` (TypeScript scripting runtime)
- `/` and click the "Python" button in the status bar (Python panel demo)

## Production/Tauri: `crossOriginIsolated` check

The packaged Tauri app **must** run with `globalThis.crossOriginIsolated === true` and `SharedArrayBuffer` available.
If this regresses, the Pyodide Worker backend breaks (and Python will fall back to the slower main-thread mode).

### Quick check (recommended)

Run the automated smoke check:

```bash
pnpm -C apps/desktop check:coi
```

This builds the production frontend + a release desktop binary and launches it in a special mode that exits with
success/failure after evaluating `globalThis.crossOriginIsolated`, `SharedArrayBuffer`, and basic Worker support inside the WebView.

If you already built the app (for example via `cargo tauri build` / `tauri-action`), you can skip rebuilding and run the check against
the existing artifacts:

```bash
pnpm -C apps/desktop check:coi -- --no-build
```

Optional: override the binary path explicitly (useful when multiple Cargo target outputs exist, e.g. `target/release/...` vs `target/<triple>/release/...`):

```bash
pnpm -C apps/desktop check:coi -- --no-build --bin <path-to-formula-desktop>
```

Optional (Linux/CI): if the app occasionally hangs in headless environments, you can tune the outer timeout used by the check:

```bash
FORMULA_COI_TIMEOUT_SECS=60 pnpm -C apps/desktop check:coi -- --no-build
```

CI note: the desktop release workflow runs this check on Linux (and, by default, macOS/Windows) **after** the Tauri build step,
reusing the already-built artifacts (`--no-build`). To temporarily skip the check on macOS/Windows (while keeping the Linux signal),
set the GitHub Actions variable `FORMULA_COI_CHECK_ALL_PLATFORMS=0` (or `false`).

### Manual check

1. Build a production desktop binary:

   ```bash
   pnpm -C apps/desktop build
   bash scripts/cargo_agent.sh build -p formula-desktop-tauri --features desktop --bin formula-desktop --release
   ```

2. Launch the built app (platform-specific binary path under `target/release/`).
3. On startup, the app will show a long-lived **error toast** if cross-origin isolation is missing.

### Manual verification (DevTools)

If you have DevTools access in the packaged WebView, run:

```js
globalThis.crossOriginIsolated
typeof SharedArrayBuffer !== "undefined"
```

## Tauri capability permission identifier check

Tauri v2 capabilities (`apps/desktop/src-tauri/capabilities/*.json`) reference permission identifiers that can drift when
Tauri core/plugins are upgraded. To validate that the capability files only reference identifiers supported by your
installed toolchain, run:

```bash
pnpm -C apps/desktop check:tauri-permissions
```

This check runs `cargo tauri permission ls` under the hood, so it requires the Tauri CLI and (on Linux) the system WebView
dependencies.

## Production/Tauri: clipboard smoke check

The packaged Tauri app should be able to round-trip key clipboard formats (at minimum `text/plain`, `text/html`, and
`image/png`) via the native clipboard backends. This helps catch OS-specific regressions in Windows CF_HTML/PNG handling,
macOS pasteboard conversions, etc.

### Quick check (opt-in)

Run the automated smoke check:

```bash
pnpm -C apps/desktop check:clipboard
```

This builds the production frontend + a release desktop binary and launches it in a special mode that writes/reads
clipboard formats and exits with:

- `0` on success
- `1` on functional failure (clipboard APIs available but incorrect)
- `2` on internal error/timeout (e.g. backend unavailable)

## Extensions / Marketplace (Tauri/WebView runtime — no Node)

Formula Desktop runs extensions inside the **WebView** runtime (no Electron-style Node integration).

At a high level:

- **Runtime:** `BrowserExtensionHost` (Web Worker-based extension host)
  - Source: `packages/extension-host/src/browser/index.mjs` (exported as `@formula/extension-host/browser`)
  - Each extension runs in its own module `Worker` (`packages/extension-host/src/browser/extension-worker.mjs`).
- **Installer + package store:** `WebExtensionManager` (IndexedDB-backed installer/loader)
  - Source: `packages/extension-marketplace/src/WebExtensionManager.ts` (exported as `@formula/extension-marketplace`)
  - Downloads signed `.fextpkg` packages from the marketplace, verifies SHA-256 + Ed25519 signatures **in the
    WebView**, and persists verified bytes + metadata to IndexedDB.
  - **Boot:** `WebExtensionManager.loadAllInstalled()` loads all installed extensions and triggers startup semantics
    (`onStartupFinished` + initial `workbookOpened`) in a way that avoids spamming already-running extensions.
- **Desktop integration:** `DesktopExtensionHostManager` (`apps/desktop/src/extensions/extensionHostManager.ts`)
  - Wires the host/manager into the desktop UI (toasts/prompts/panels, base URL config) and loads built-in +
    IndexedDB-installed extensions on demand.

Extension panels (`contributes.panels` / `formula.ui.createPanel`) are rendered in a sandboxed `<iframe>` with a
 restrictive CSP (see `apps/desktop/src/extensions/ExtensionPanelBody.tsx`), so panel HTML cannot load remote scripts,
 cannot make network requests directly, and cannot run inline `<script>` blocks (scripts must be loaded via `data:`/`blob:`
 URLs). The desktop also scrubs Tauri IPC globals (`__TAURI__`, etc) from the iframe as a defense-in-depth measure (see
 `window.__formulaWebviewSandbox` marker inside the iframe). Panels should communicate with extension code via
 `postMessage`.

The extension worker runtime also locks down Tauri globals (`__TAURI__`, `__TAURI_IPC__`, etc) before loading extension
modules (defense-in-depth so untrusted extension code can't call native commands directly).

### Where extensions live / what persists

**Installed packages (code):**

- Stored in **IndexedDB** under database `formula.webExtensions` (see `WebExtensionManager`):
  - `installed` store: `{ id, version, installedAt }`
  - `packages` store: `{ key: "${id}@${version}", bytes, verified }`

**Permission grants (user decisions):**

- Stored by `BrowserExtensionHost` in `localStorage["formula.extensionHost.permissions"]`
  (per-extension record of granted permissions).
  - Note: the key is removed entirely when the normalized store is empty or corrupted (clean slate).

**Extension storage + config (`formula.storage` / `formula.config` APIs):**

- Stored in `localStorage` via `LocalStorageExtensionStorage`:
  - key prefix: `formula.extensionHost.storage.`
  - key per extension: `formula.extensionHost.storage.<extensionId>`
  - Note: per-extension keys are removed entirely when the store is empty or corrupted (clean slate).

**Panel layout seed data (contributed panels):**

- Stored in a synchronous localStorage “seed store” so contributed panel ids can be registered *before* layout
  deserialization on startup:
  - `localStorage["formula.extensions.contributedPanels.v1"]`
  - Note: the key is removed entirely when the normalized store is empty or corrupted (clean slate).

To “reset” extensions in dev, clear:

- IndexedDB database `formula.webExtensions`
- localStorage keys `formula.extensionHost.permissions` and `formula.extensionHost.storage.*`
- localStorage key `formula.extensions.contributedPanels.v1` (optional; only affects panel layout persistence)

You can also do a quick reset from the console:

```js
indexedDB.deleteDatabase("formula.webExtensions");
localStorage.removeItem("formula.extensionHost.permissions");
localStorage.removeItem("formula.extensions.contributedPanels.v1");
// (Optional) clear per-extension storage keys under formula.extensionHost.storage.*
```

### Debugging

- Use WebView DevTools to inspect:
  - **Application → IndexedDB →** `formula.webExtensions` (installed packages + versions)
  - **Application → Local Storage →** `formula.extensionHost.*` (permissions + extension storage)
  - **Application → Local Storage →** `formula.extensions.contributedPanels.v1` (contributed panel seed store)
  - **Sources → Workers** to debug the extension worker runtime (`extension-worker.mjs`)
  - `window.__formulaExtensionHost` / `window.__formulaExtensionHostManager` (debug-only globals exposed for e2e tests)

To open the built-in Extensions panel (and trigger the lazy extension host boot), use the ribbon:
**Home → Panels → Extensions**.

### Marketplace base URL (Desktop)

The marketplace base URL is chosen by `getMarketplaceBaseUrl()` (`apps/desktop/src/panels/marketplace/getMarketplaceBaseUrl.ts`).

You can override it in two ways:

- **Runtime (DevTools)**: `localStorage["formula:marketplace:baseUrl"]` and reload.
- **Build/runtime config**: set `VITE_FORMULA_MARKETPLACE_BASE_URL` (used by Vite / packaged desktop builds).

For the `localStorage` override, you can provide either:

- an **origin** (`https://marketplace.formula.app`) — it will be normalized to `.../api`, or
- the explicit API base URL (`https://marketplace.formula.app/api`).

```js
localStorage.setItem("formula:marketplace:baseUrl", "https://marketplace.formula.app/api");
location.reload();
```

Example with an env var (for a local marketplace server):

```bash
VITE_FORMULA_MARKETPLACE_BASE_URL=http://127.0.0.1:8787 pnpm -C apps/desktop dev
```

For running a local marketplace server (and registering a publisher for test publishes), see
`services/marketplace/README.md`.

To open the built-in Marketplace panel:

- Use the ribbon: **View → Panels → Marketplace**
- Or run from DevTools:

```js
window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "marketplace" } }));
```

### Legacy Node-only installer/runtime (deprecated)

The repo still contains Node-only marketplace/host modules used by Node integration tests and earlier experiments:

- Installer: `apps/desktop/tools/marketplace/extensionManager.js`
- Marketplace client: `apps/desktop/tools/marketplace/client.js`
- Runtime: `apps/desktop/tools/extensions/ExtensionHostManager.js`

They rely on `node:fs` / `worker_threads` and are **not used by the desktop renderer**.

CI enforces that the desktop renderer (`apps/desktop/src/**`) stays Node-free (no `node:*`/`fs`/`path` imports, and no
imports from `apps/desktop/tools/**` or `apps/desktop/scripts/**`). If you need Node-only code, keep it under one of
those tooling directories and bridge it into the real app via IPC/Tauri plumbing.

See `docs/10-extensibility.md` for the end-to-end flow.

## Real-time collaboration (Yjs)

The desktop dev server can run in a real-time collaborative mode backed by the local Yjs sync server (`services/sync-server`).
Collaboration is enabled/configured via URL query params.

### 1) Start the sync server

From the repo root:

```bash
pnpm dev:sync
```

Equivalent:

```bash
pnpm -C services/sync-server dev
```

Notes:

- Default WebSocket URL: `ws://127.0.0.1:1234`
- The sync server requires auth. In non-production environments it defaults to the shared dev token: `dev-token`
  (you can override with `SYNC_SERVER_AUTH_TOKEN=...`).

### 2) Start the desktop app (Vite dev server)

```bash
pnpm -C apps/desktop dev
```

This serves the app at `http://localhost:4174`.

### 3) Open two clients to the same doc (URL params)

Each desktop client gets a **stable collaboration identity** (`id`/`name`/`color`) persisted in localStorage
(`formula:collab:user`). For local testing you can override the identity per-window via URL query params.

Open two browser windows/tabs with the same `docId`, but different user identities:

```text
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=my-doc&token=dev-token&collabUserId=u1&collabUserName=User%201&collabUserColor=%23ff0000
http://localhost:4174/?collab=1&wsUrl=ws://127.0.0.1:1234&docId=my-doc&token=dev-token&collabUserId=u2&collabUserName=User%202&collabUserColor=%2300ff00
```

Params:

- `collab=1` enables collaboration mode
- `wsUrl` is the base sync-server URL (no trailing `/docId`)
- `docId` is the shared document/room name (must match across clients)
- `token` must match the sync-server auth token (defaults to `dev-token`)
- `collabEncrypt=1` enables **dev-only** end-to-end cell encryption for a deterministic demo range.
  - Use this to exercise encrypted cell payloads (`enc`) end-to-end across multiple clients.
  - To verify masking, open one client with `collabEncrypt=1` and another without it.
  - The key is derived deterministically from `docId` + a hardcoded dev salt (testing only; not production key management).
- `collabEncryptRange=Sheet1!A1:C10` optionally overrides the encrypted range (default: `Sheet1!A1:C10`).
  - The sheet part uses the same syntax as formulas (sheet *name*). When a sheet-name resolver is available, the desktop
    will map that name to the stable sheet id used by collab cell keys (so this works even when ids differ from names).
- `collabPersistence=0` disables local persistence (IndexedDB) for debugging/tests (default: enabled)
  - Legacy alias (deprecated): `collabOffline=0`
- `collabUserId`, `collabUserName`, `collabUserColor` override the per-client identity used for presence/comments/conflicts
  - `collabUserColor` must be URL-encoded (`#` → `%23`)
  - Legacy aliases are still accepted: `userId`, `userName`, `userColor`

### 4) Expected behavior

With two clients connected to the same `docId` you should see:

- Cell edits sync between windows/tabs
- Cell edits persist across reloads/crash via IndexedDB (offline-first); on reconnect, Yjs merges offline edits into the shared doc
- Comments sync between windows/tabs
- Presence (remote cursors / selections) rendered in the grid
- A conflicts UI when two clients concurrently edit the same cell

## Power Query caching + credentials (security)

The desktop app persists some Power Query state across restarts:

- Query-result cache (IndexedDB)
- Connector credentials (OS keychain-backed encrypted store)
- Refresh scheduling state

All of this is encrypted-at-rest in production builds:

- Query-result caching is wrapped in `EncryptedCacheStore` (AES-256-GCM). The 32-byte
  cache key is generated once and stored in the OS keychain via a Tauri command.
- Credential + refresh-state stores are encrypted blobs on disk with key material in
  the OS keychain (Rust/Tauri storage layer).

## OAuth redirects (desktop)

Some Power Query connectors use OAuth (typically auth-code + PKCE). Formula Desktop can capture OAuth redirects via:

- **Deep link (preferred):** `formula://oauth/callback` (custom URI scheme handled by the OS)
- **RFC 8252 loopback (fallback):** `http://127.0.0.1:<port>/oauth/callback` (also supports `localhost` / `[::1]`)

Redirects observed by the Rust host are forwarded to the UI via the `oauth-redirect` event. On startup the UI emits
`oauth-redirect-ready` to flush any redirects that arrived before the listener was installed.

Details: `docs/11-desktop-shell.md` → “OAuth redirects (`formula://…` deep links vs RFC 8252 loopback)”.

## AI

Formula is a **Cursor product**: all AI features are powered by Cursor's backend—no local models, no API keys, no provider configuration.

### Formula bar tab completion

The formula bar supports tab completion suggestions while you type. Suggestions combine:

- fast rule-based suggestions (function names, ranges, named ranges, sheet-qualified ranges, tables/structured refs)
- optional Cursor backend completions (Cursor-managed; no user API keys, no provider/model selection)
- optional inline preview values (when the lightweight evaluator supports the suggested formula)

Backend completions are **time-bounded** (defaults to a ~100ms budget) and ignored on timeout/cancel so the formula bar stays responsive. If the backend is unavailable, the UI continues to show rule-based suggestions.

Backend completions are routed to a Cursor-managed endpoint (e.g. `/api/ai/tab-completion`) and are Cursor/build-managed (not a user-facing setting).

Notes:

- Structured-reference preview is evaluated for structured refs that resolve to a rectangular range when table range metadata is available (e.g. `Table[Column]`, `Table[[#All],[Column]]`, and contiguous multi-column forms like `Table[[#All],[Col1],[Col2]]` / `Table[[Col1]:[Col3]]`). Non-rectangular unions still fall back to `(preview unavailable)`.
