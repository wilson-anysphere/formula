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
