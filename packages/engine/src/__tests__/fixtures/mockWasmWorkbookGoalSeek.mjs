// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`.
//
// This is loaded by `packages/engine/src/engine.worker.ts` via dynamic import (runtime string),
// so it must be plain JS (not TS) to avoid relying on Vite/Vitest transforms.

export default async function init() {
  // No-op.
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

  goalSeek(request) {
    // Return either the new `{ result, changes }` payload shape or the legacy flat shape
    // depending on whether the caller supplies `derivativeStep` (used here as a simple
    // test toggle).
    if (request && request.derivativeStep != null) {
        return {
          result: {
          status: "  Converged  ",
          solution: 5,
          iterations: 3,
          finalOutput: 25,
          finalError: 0,
        },
        changes: [
          { sheet: "Sheet1", address: "A1", value: 5 },
          { sheet: "Sheet1", address: "B1", value: 25 },
          // Simulate wasm-bindgen `Option<T>` -> `undefined` mapping for blanks.
          { sheet: "Sheet1", address: "C1", value: undefined },
        ],
      };
    }

    return {
      success: true,
      status: "  Converged  ",
      solution: 5,
      iterations: 3,
      finalError: 0,
      // Intentionally omit `finalOutput` to exercise the worker's legacy compatibility shim
      // (it should compute `finalOutput` from `targetValue + finalError`).
    };
  }
}

// Editor-tooling exports (unused by the goalSeek-focused worker tests, but included to keep the
// mock module closer to real formula-wasm shapes).
export function lexFormula() {
  return [];
}
