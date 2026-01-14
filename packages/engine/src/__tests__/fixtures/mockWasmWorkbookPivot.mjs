// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`.
//
// This is loaded by `packages/engine/src/engine.worker.ts` via dynamic import (runtime string),
// so it must be plain JS (not TS) to avoid relying on Vite/Vitest transforms.

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
  setRange(_range, _values, _sheet) {}
  recalculate(_sheet) {
    return [];
  }

  calculatePivot(sheet, _sourceRangeA1, _destinationTopLeftA1, _config) {
    recordCall("calculatePivot", sheet);
    return {
      writes: [
        // Simulate wasm-bindgen `Option<T>` -> `undefined` mapping for blanks.
        { sheet: "Sheet1", address: "D1", value: undefined },
        { sheet: "Sheet1", address: "E1", value: 123 },
      ],
    };
  }
}
