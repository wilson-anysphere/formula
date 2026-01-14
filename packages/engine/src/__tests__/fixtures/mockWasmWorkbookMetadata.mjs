// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`.
//
// This is loaded by `packages/engine/src/engine.worker.ts` via dynamic import (runtime string),
// so it must be plain JS (not TS) to avoid relying on Vite/Vitest transforms.
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

  // Minimal surface required by the worker's type expectations.
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
  setCells(updates) {
    recordCall("setCells", updates);
  }
  setCellRich(address, value, sheet) {
    recordCall("setCellRich", address, value, sheet);
  }
  setRange(range, values, sheet) {
    recordCall("setRange", range, values, sheet);
  }
  applyOperation(op) {
    recordCall("applyOperation", op);
    return { changedCells: [], movedRanges: [], formulaRewrites: [] };
  }
  recalculate(_sheet) {
    return [];
  }

  setSheetDimensions(sheet, rows, cols) {
    recordCall("setSheetDimensions", sheet, rows, cols);
  }

  getSheetDimensions(sheet) {
    recordCall("getSheetDimensions", sheet);
    return { rows: 100, cols: 200 };
  }

  renameSheet(oldName, newName) {
    recordCall("renameSheet", oldName, newName);
    return true;
  }

  setWorkbookFileMetadata(directory, filename) {
    recordCall("setWorkbookFileMetadata", directory, filename);
  }

  setInfoOriginForSheet(sheet, origin) {
    recordCall("setInfoOriginForSheet", sheet, origin);
  }

  setCellStyleId(sheet, address, styleId) {
    recordCall("setCellStyleId", sheet, address, styleId);
  }

  setRowStyleId(sheet, row, styleId) {
    recordCall("setRowStyleId", sheet, row, styleId);
  }

  setColStyleId(sheet, col, styleId) {
    recordCall("setColStyleId", sheet, col, styleId);
  }

  setFormatRunsByCol(sheet, col, runs) {
    recordCall("setFormatRunsByCol", sheet, col, runs);
  }

  setSheetDefaultStyleId(sheet, styleId) {
    recordCall("setSheetDefaultStyleId", sheet, styleId);
  }

  setColWidth(sheet, col, width) {
    recordCall("setColWidth", sheet, col, width);
  }

  setColWidthChars(sheet, col, widthChars) {
    recordCall("setColWidthChars", sheet, col, widthChars);
  }

  setSheetDisplayName(sheetId, name) {
    recordCall("setSheetDisplayName", sheetId, name);
  }

  setColHidden(sheet, col, hidden) {
    recordCall("setColHidden", sheet, col, hidden);
  }

  setColFormatRuns(sheet, col, runs) {
    recordCall("setColFormatRuns", sheet, col, runs);
  }

  internStyle(style) {
    recordCall("internStyle", style);
    return 42;
  }

  static fromJson(_json) {
    return new WasmWorkbook();
  }

  static fromXlsxBytes(_bytes) {
    return new WasmWorkbook();
  }
}

// Editor-tooling exports (unused by the metadata-focused tests, but included to keep the mock
// module closer to real formula-wasm shapes).
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
