# @formula/python-runtime

Modern Python scripting for Formula with two runtimes:

- **`NativePythonRuntime`** (desktop / Node): spawns a local Python interpreter and runs user scripts with a JSON-RPC bridge over stdio.
- **`PyodideRuntime`** (web / webview): loads Pyodide in a dedicated Worker, installs the in-repo `formula` Python package, and bridges spreadsheet operations via an injected JS module.

This package is intentionally lightweight and bridge-driven: the host application provides the spreadsheet API implementation.

## Usage

### Native Python (desktop / Node)

```js
import { NativePythonRuntime, MockWorkbook } from "@formula/python-runtime";

const workbook = new MockWorkbook(); // example bridge (tests)
const runtime = new NativePythonRuntime();

await runtime.execute(
  `
import formula
sheet = formula.active_sheet
sheet["A1"] = 42
sheet["A2"] = "=A1*2"
`,
  { api: workbook },
);
```

`execute()` resolves to an object with these additional fields:
- `stdout: string` (always `""` for `NativePythonRuntime` — stdout is reserved for the JSON protocol stream)
- `stderr: string` (captured user output; note that native Python redirects `sys.stdout` to `sys.stderr`)

If execution fails, the thrown `Error` also includes `err.stdout` / `err.stderr` (when available).

If you already use `apps/desktop`'s `DocumentController`, you can use the
included adapter:

```js
import { DocumentController } from "../../apps/desktop/src/document/documentController.js";
import { DocumentControllerBridge, NativePythonRuntime } from "@formula/python-runtime";

const doc = new DocumentController();
const api = new DocumentControllerBridge(doc);
const runtime = new NativePythonRuntime();

await runtime.execute(`import formula\nformula.active_sheet["A1"] = 1\n`, { api });
```

### Pyodide (web / webview)

```js
import { PyodideRuntime } from "@formula/python-runtime";

const runtime = new PyodideRuntime({
  api: mySpreadsheetBridge,
  // Optional: choose a backend mode ("auto" | "worker" | "mainThread").
  // - auto (default): prefers worker mode when COOP/COEP + SharedArrayBuffer are available.
  // - worker: force worker mode (requires crossOriginIsolated + SharedArrayBuffer).
  // - mainThread: force main-thread Pyodide (UI will block while scripts run).
  // mode: "auto",
  //
  // Recommended to self-host Pyodide assets (especially for crossOriginIsolated environments):
  // indexURL: "/pyodide/v0.25.1/full/",
});

await runtime.initialize({
  // Optional overrides:
  // mode: "auto",
  // permissions: { filesystem: "none", network: "none" },
  // rpcTimeoutMs: 5000,
});

await runtime.execute(`
import formula
formula.active_sheet["A1"] = 123
`);
```

### Backends: Worker vs main-thread

The Pyodide runtime supports two internal backends:

- **Worker backend (preferred)**: loads Pyodide in a Worker and keeps the Python `formula` API synchronous via a
  `SharedArrayBuffer + Atomics` RPC bridge between the Worker (Pyodide) and the host.
- **Main-thread backend (fallback)**: loads Pyodide on the main thread and calls the host spreadsheet bridge
  synchronously. This works in non-COOP/COEP contexts but will block the UI while Python runs.

That typically requires `crossOriginIsolated` in browsers, which means serving
your app with COOP/COEP headers:

- `Cross-Origin-Opener-Policy: same-origin`
- `Cross-Origin-Embedder-Policy: require-corp` (or `credentialless`)

In `mode: "auto"` (default), the runtime selects the Worker backend when possible and otherwise falls back to the
main-thread backend.

Notes:
- Main-thread mode requires the spreadsheet bridge to be synchronous (methods must not return Promises).
- Timeouts/interrupts are best-effort in main-thread mode (the UI may freeze during long-running scripts).

For the `apps/desktop` Vite webview in this repository:

- `apps/desktop/vite.config.ts` sets these headers for dev/preview servers.
- `apps/desktop/scripts/ensure-pyodide-assets.mjs` downloads Pyodide
  `v0.25.1/full/*` into `apps/desktop/public/pyodide/v0.25.1/full/` so Pyodide
  can be loaded from the same origin (required under COEP).

## Host spreadsheet bridge contract

The host application must provide an `api` object implementing (or dispatching) these RPC methods:

- `get_active_sheet_id() -> str`
- `get_sheet_id({ name }) -> str | null`
- `create_sheet({ name, index? }) -> str`
- `get_sheet_name({ sheet_id }) -> str`
- `rename_sheet({ sheet_id, name }) -> null`
- `get_range_values({ range }) -> any[][]`
- `set_range_values({ range, values }) -> null`
- `set_cell_value({ range, value }) -> null`
- `get_cell_formula({ range }) -> str | null`
- `set_cell_formula({ range, formula }) -> null`
- `clear_range({ range }) -> null`

`range` is a JSON object:

```json
{
  "sheet_id": "Sheet1",
  "start_row": 0,
  "start_col": 0,
  "end_row": 1,
  "end_col": 0
}
```

## Permissions

Permissions are best-effort guardrails (not a hardened security boundary).

```ts
type Permissions = {
  filesystem?: "none" | "read" | "readwrite";
  // legacy casing supported by the Python sandbox:
  fileSystem?: "none" | "read" | "readwrite";

  network?: "none" | "allowlist" | "full";
  networkAllowlist?: string[];
};
```

Notes:
- For native Python, `"allowlist"` is enforced by wrapping `socket.create_connection` / `socket.socket.connect`.
- For Pyodide, `"allowlist"` is enforced by wrapping `fetch`/`WebSocket` in the worker or main-thread runtime (best-effort).

## Bundled Python files (Pyodide)

The in-repo `python/formula_api/**` package is bundled into
`src/formula-files.generated.js` for installation into Pyodide's virtual
filesystem.

Regenerate it after editing the Python package:

```bash
node packages/python-runtime/scripts/generate-formula-files.js
```
