// Minimal ESM module that emulates an older wasm-bindgen build of `crates/formula-wasm` where
// `setCellStyleId` used a sheet-last signature: `setCellStyleId(address, styleId, sheet?)`.
//
// The modern engine worker prefers the sheet-first signature, but should fall back to this legacy
// shape when it detects the mismatch.
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

function colNameToIndex(colName) {
  let n = 0;
  for (const ch of colName.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`Invalid column name: ${colName}`);
    }
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

function assertValidA1(address) {
  const trimmed = String(address ?? "").trim();
  const match = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(trimmed);
  if (!match) {
    throw `invalid cell address: ${address}`;
  }
  const [, colName] = match;
  const col0 = colNameToIndex(colName);
  if (col0 < 0 || col0 >= 16_384) {
    throw `invalid cell address: ${address}`;
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
  setRange(_range, _values, _sheet) {}
  recalculate(_sheet) {
    return [];
  }

  // Legacy signature: (address, styleId, sheet?)
  setCellStyleId(address, styleId, sheet) {
    // Validate the address similarly to formula-wasm so the worker's signature-detection fallback
    // can trigger when it accidentally sends the sheet name as the address.
    assertValidA1(address);
    recordCall("setCellStyleId", address, styleId, sheet);
  }

  static fromJson(_json) {
    return new WasmWorkbook();
  }

  static fromXlsxBytes(_bytes) {
    return new WasmWorkbook();
  }
}

// Editor-tooling exports (unused by the metadata-focused tests).
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

