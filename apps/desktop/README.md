# Formula Desktop (Vite/Webview)

## Pyodide / Python scripting

The Pyodide-based Python runtime (`@formula/python-runtime`) uses
`SharedArrayBuffer + Atomics` for synchronous spreadsheet RPC. In browsers and
webviews this requires a **cross-origin isolated** context.

This app’s Vite dev/preview servers are configured to enable that:

- `Cross-Origin-Opener-Policy: same-origin`
- `Cross-Origin-Embedder-Policy: require-corp`

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

## Extensions / Marketplace (current status)

The repo contains a complete **extension runtime + marketplace installer** implementation, but it currently lives in
**Node-only modules**:

- Installer: `apps/desktop/src/marketplace/extensionManager.js`
- Runtime: `apps/desktop/src/extensions/ExtensionHostManager.js` (wraps `packages/extension-host`)

These modules use `node:fs` and `worker_threads` and are **not wired into the Vite/WebView runtime yet**. They are
used by Node integration tests and are intended to be bridged into the real desktop app via IPC/Tauri plumbing.

See `docs/10-extensibility.md` for the end-to-end flow and hot-reload behavior.
