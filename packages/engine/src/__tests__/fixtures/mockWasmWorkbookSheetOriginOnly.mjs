// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`,
// exposing only `setSheetOrigin` (and intentionally omitting the legacy
// `setInfoOriginForSheet`) so worker tests can exercise fallback logic.
//
// Tests record calls in `globalThis.__ENGINE_WORKER_TEST_CALLS__`.

export default async function init() {
  // No-op.
}

function recordCall(name, ...args) {
  const calls = globalThis.__ENGINE_WORKER_TEST_CALLS__;
  if (Array.isArray(calls)) {
    calls.push([name, ...args]);
  }
}

export class WasmWorkbook {
  constructor() {}

  toJson() {
    return "{}";
  }

  recalculate(_sheet) {
    return [];
  }

  setSheetOrigin(sheet, origin) {
    recordCall("setSheetOrigin", sheet, origin);
  }

  static fromJson(_json) {
    return new WasmWorkbook();
  }
}

// Editor-tooling exports (unused by these tests, but included to keep the mock module closer to
// real formula-wasm shapes).
export function lexFormula() {
  return [];
}
export function lexFormulaPartial() {
  return { tokens: [], error: null };
}
export function parseFormulaPartial() {
  return { ast: null, error: null, context: { function: null } };
}

