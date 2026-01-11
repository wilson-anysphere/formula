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

## Formula bar AI tab-completion (local model)

The formula bar has an “AI-native” tab completion layer that combines:

- fast rule-based suggestions (function names, ranges, named ranges, sheet-qualified ranges, tables/structured refs)
- **optional** local-model suggestions via [Ollama](https://ollama.com/)
- optional inline preview values (when the lightweight evaluator supports the suggested formula)

### Enabling the local model

Local model completions are controlled via `localStorage` flags (use DevTools in the WebView):

- `formula:aiCompletion:localModelEnabled` = `true`
- `formula:aiCompletion:localModelName` = model name (default: `formula-completion`)
- `formula:aiCompletion:localModelBaseUrl` = Ollama base URL (default: `http://localhost:11434`)

Notes:

- Completions are time-bounded (defaults to a ~200ms budget) so the formula bar stays responsive even if Ollama is slow/unavailable.
- Structured-reference preview is currently not evaluated (the UI will show `(preview unavailable)` for those suggestions).
