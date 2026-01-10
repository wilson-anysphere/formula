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

### Pyodide (web / webview)

```js
import { PyodideRuntime } from "@formula/python-runtime";

const runtime = new PyodideRuntime({
  api: mySpreadsheetBridge,
  // Recommended to self-host Pyodide assets for crossOriginIsolated environments:
  // indexURL: "/pyodide/v0.25.1/full/",
});

await runtime.initialize({
  // Optional overrides:
  // permissions: { filesystem: "none", network: "none" },
  // rpcTimeoutMs: 5000,
});

await runtime.execute(`
import formula
formula.active_sheet["A1"] = 123
`);
```

### Worker / SharedArrayBuffer requirement

The Pyodide runtime keeps the Python `formula` API synchronous by using a
SharedArrayBuffer + Atomics-based RPC between the Worker (Pyodide) and the host.

That typically requires `crossOriginIsolated` in browsers. If SharedArrayBuffer
is not available, scripts can still run, but calls into `formula` will raise.

## Host spreadsheet bridge contract

The host application must provide an `api` object implementing (or dispatching) these RPC methods:

- `get_active_sheet_id() -> str`
- `get_sheet_id({ name }) -> str | null`
- `create_sheet({ name }) -> str`
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
- Native allowlist enforcement is not implemented (yet); `"allowlist"` behaves like `"full"` for native Python.
- For Pyodide, `"allowlist"` is enforced by wrapping `fetch`/`WebSocket` in the worker.

## Bundled Python files (Pyodide)

The in-repo `python/formula_api/**` package is bundled into
`src/formula-files.generated.js` for installation into Pyodide's virtual
filesystem.

Regenerate it after editing the Python package:

```bash
node packages/python-runtime/scripts/generate-formula-files.js
```

