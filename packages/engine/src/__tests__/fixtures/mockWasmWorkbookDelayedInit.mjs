// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`, but with an
// intentionally delayed `init()` so tests can simulate in-flight requests while the worker is
// being re-initialized.
//
// Tests control the delay by setting `globalThis.__ENGINE_WORKER_DELAY_INIT_PROMISE__` to a
// Promise that resolves when initialization should continue.

export default async function init() {
  const gate = globalThis.__ENGINE_WORKER_DELAY_INIT_PROMISE__;
  if (gate && typeof gate.then === "function") {
    await gate;
  }
}

export class WasmWorkbook {
  toJson() {
    return "{}";
  }
  getCell(address, sheet) {
    return { sheet: sheet ?? "Sheet1", address, input: null, value: null };
  }
  getRange() {
    return [];
  }
  setCell() {}
  setRange() {}
  recalculate() {
    return [];
  }

  static fromJson() {
    return new WasmWorkbook();
  }
}

// Editor-tooling exports (unused by these tests, but included for parity).
export function lexFormula() {
  return [];
}
export function lexFormulaPartial() {
  return { tokens: [], error: null };
}
export function parseFormulaPartial() {
  return { ast: null, error: null, context: { function: null } };
}

