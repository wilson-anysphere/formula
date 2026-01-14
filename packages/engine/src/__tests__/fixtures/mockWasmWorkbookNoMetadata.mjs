// Mock wasm module that intentionally omits the workbook metadata setter methods.
// Used to verify that `engine.worker.ts` surfaces clear errors when a WASM build
// doesn't support newer APIs.

export default async function init() {}

export class WasmWorkbook {
  constructor() {}

  toJson() {
    return "{}";
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
export function parseFormulaPartial() {
  return { ast: null, error: null, context: { function: null } };
}

