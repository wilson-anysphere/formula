# Formula Desktop (Vite/Webview)

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

Additionally, Pyodide assets are self-hosted under the same origin at:

`/pyodide/v0.25.1/full/`

Running `pnpm -C apps/desktop dev` (or `pnpm -C apps/desktop build`) will
download the required Pyodide files into `apps/desktop/public/pyodide/` via
`scripts/ensure-pyodide-assets.mjs`.

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
success/failure after evaluating `globalThis.crossOriginIsolated` and `SharedArrayBuffer` inside the WebView.

### Manual check

1. Build a production desktop binary:

   ```bash
   pnpm -C apps/desktop build
   cd apps/desktop
   bash ../../scripts/cargo_agent.sh tauri build --no-bundle
   ```

2. Launch the built app (platform-specific binary path under `target/release/`).
3. On startup, the app will show a long-lived **error toast** if cross-origin isolation is missing.

### Manual verification (DevTools)

If you have DevTools access in the packaged WebView, run:

```js
globalThis.crossOriginIsolated
typeof SharedArrayBuffer !== "undefined"
```

## Extensions / Marketplace (current status)

The desktop app’s canonical extension runtime is **no-Node** and runs entirely inside the WebView:

- **Runtime:** `BrowserExtensionHost` (runs each extension in a module `Worker`)
- **Installer + persistence:** `WebExtensionManager` (downloads + verifies signed v2 `.fextpkg` packages and stores them
  in IndexedDB)

Marketplace installs are loaded via:
`WebExtensionManager.loadInstalled(...)` → `BrowserExtensionHost.loadExtension(...)` using a `blob:`/`data:` module URL
(no filesystem extraction required).

### Where extensions are stored (and how to clear for dev)

Installed packages + metadata are stored in IndexedDB database **`formula.webExtensions`**.

To reset installed extensions during development:

- DevTools → **Application** → **IndexedDB** → delete `formula.webExtensions`, then reload, or
- run in the console: `indexedDB.deleteDatabase("formula.webExtensions")`

Note: permission grants and per-extension storage are persisted separately in `localStorage` (keys under
`formula.extensionHost.*`).

### Legacy Node-only implementation (tests/legacy tooling)

These Node-only modules are kept for legacy tooling + integration tests and are **not** used by the Tauri/WebView
runtime:

- `apps/desktop/src/marketplace/extensionManager.js`
- `apps/desktop/src/extensions/ExtensionHostManager.js`

See `docs/10-extensibility.md` for the end-to-end flow.

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

- Structured-reference preview is evaluated for simple table column refs (`Table[Column]` / `Table[[#All],[Column]]`) when table range metadata is available. More complex structured refs still fall back to `(preview unavailable)`.
