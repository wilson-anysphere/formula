// Minimal ESM module that emulates an older wasm-bindgen build of `crates/formula-wasm`
// that does NOT export `getWorkbookInfo()`.
//
// `engine.worker.ts` should fall back to parsing `toJson()` for getWorkbookInfo requests.
// This fixture encodes a `sheetOrder` that differs from the object key order so tests can
// verify the worker respects `sheetOrder`.

export default async function init() {
  // No-op.
}

export class WasmWorkbook {
  constructor() {}

  // Minimal surface required by the worker's type expectations.
  toJson() {
    return JSON.stringify({
      sheetOrder: ["Sheet2", "Sheet1", "Empty"],
      sheets: {
        // Deliberately out of order to ensure consumers don't rely on Object.keys() ordering.
        Empty: { cells: {} },
        Sheet1: { cells: { A1: 1 } },
        Sheet2: { cells: { B2: 2 } }
      }
    });
  }

  getCell(address, sheet) {
    return { sheet: sheet ?? "Sheet1", address, input: null, value: null };
  }
  getRange(_range, _sheet) {
    return [];
  }
  setCell(_address, _value, _sheet) {}
  setRange(_range, _values, _sheet) {}
  recalculate(_sheet) {
    return [];
  }

  static fromJson(_json) {
    return new WasmWorkbook();
  }
}

export function lexFormula() {
  return [];
}
export function lexFormulaPartial() {
  return { tokens: [], error: null };
}
export function parseFormulaPartial() {
  return { ast: null, error: null, context: { function: null } };
}
export function rewriteFormulasForCopyDelta({ requests }) {
  return (requests ?? []).map((r) => r.formula ?? "");
}

