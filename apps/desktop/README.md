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

For packaged (production) Tauri builds, the same headers must also be added to the
Tauri asset/custom protocol responses — see `docs/11-desktop-shell.md` (cross-origin isolation section).

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

1. Build a production desktop binary:

   ```bash
   pnpm -C apps/desktop build
   cd apps/desktop/src-tauri
   cargo tauri build --no-bundle
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

The repo contains a complete **extension runtime + marketplace installer** implementation, but it currently lives in
**Node-only modules**:

- Installer: `apps/desktop/src/marketplace/extensionManager.js`
- Runtime: `apps/desktop/src/extensions/ExtensionHostManager.js` (wraps `packages/extension-host`)

These modules use `node:fs` and `worker_threads` and are **not wired into the Vite/WebView runtime yet**. They are
used by Node integration tests and are intended to be bridged into the real desktop app via IPC/Tauri plumbing.

See `docs/10-extensibility.md` for the end-to-end flow and hot-reload behavior.

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

Formula is a **Cursor product**: AI features are powered by Cursor’s backend (no local models, no API keys, no provider configuration).

The formula bar’s tab-completion includes fast heuristic suggestions (functions, ranges, named ranges, etc.). Any AI-driven completions are backend-driven.
