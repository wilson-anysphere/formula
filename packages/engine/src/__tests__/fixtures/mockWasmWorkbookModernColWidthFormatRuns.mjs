// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`,
// exposing only the modern column-width + format-run metadata APIs so worker tests can
// exercise fallback behavior.
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

  setColWidthChars(sheet, col, widthChars) {
    recordCall("setColWidthChars", sheet, col, widthChars);
  }

  setFormatRunsByCol(sheet, col, runs) {
    recordCall("setFormatRunsByCol", sheet, col, runs);
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

