# @formula/scripting

Browser-capable TypeScript scripting runtime for Formula.

This package is intentionally split into environment-specific entrypoints:

- `@formula/scripting/node` – Node `worker_threads` runtime (used by `node:test` and CI)
- `@formula/scripting/web` – WebWorker runtime (used by Vite/Tauri webviews)
- `@formula/scripting` – shared utilities (`Workbook` model, A1 helpers, `FORMULA_API_DTS`, etc.)

## Running scripts

### Recommended: module-style scripts

```ts
export default async function main(ctx: ScriptContext) {
  const values = await ctx.activeSheet.getRange("A1:B1").getValues();
  await ctx.activeSheet.getRange("C1").setValue(Number(values[0][0]) + Number(values[0][1]));
}
```

> Note: runtime imports are not supported yet (including `import ... from "..."` and dynamic `import(...)`).
> Prefer using the global script types (via `FORMULA_API_DTS`) instead of importing.

### Legacy: script-body form (top-level await)

Script bodies without `export default` are wrapped in an async function so they can use `await`:

```ts
const name = await ctx.workbook.getActiveSheetName();
ctx.ui.log("active sheet:", name);
```

## Permissions (minimal)

`ScriptRuntime.run(code, { permissions })` supports:

- `network: "none" | "allowlist" | "full"` (default: `"none"`)
- `networkAllowlist?: string[]` (hostnames)

The worker enforces permissions by wrapping `fetch` + `WebSocket` (similar to the Python/Pyodide worker).

## UI helpers

Scripts can call:

- `ctx.alert(message)`
- `ctx.confirm(message)`
- `ctx.prompt(message, defaultValue?)`

These are forwarded via RPC to the host. The web runtime uses `window.alert/confirm/prompt`; the node runtime currently throws a “not available” error for these methods.

## Monaco / editor typings

The `FORMULA_API_DTS` export contains a string version of `packages/scripting/formula.d.ts`.
Desktop can register it as an extra lib:

```js
monaco.languages.typescript.typescriptDefaults.addExtraLib(FORMULA_API_DTS, "file:///formula.d.ts");
```

## Formatting

`Range.setFormat(null)` clears formatting (equivalent to removing all style keys).
