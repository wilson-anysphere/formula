// ESM wrapper around the wasm-pack `--target nodejs` build of `crates/formula-wasm`.
//
// `packages/engine/src/engine.worker.ts` expects wasm-bindgen modules to export a default
// async init function (the `--target bundler` shape), but our vitest environment builds
// a Node-compatible CommonJS bundle in `crates/formula-wasm/pkg-node/`.
//
// This wrapper adapts the Node bundle to the worker's expectation so we can run true
// end-to-end worker RPC tests against the real WASM engine.

// ESM import of a CommonJS module yields its exports on `default`.
import nodeBundle from "../../../../../crates/formula-wasm/pkg-node/formula_wasm.js";

const wasm = nodeBundle?.default ?? nodeBundle;

export default async function init() {
  // wasm-pack `--target nodejs` initializes at import time; nothing to do here.
}

// Re-export the surface expected by `engine.worker.ts`.
export const WasmWorkbook = wasm.WasmWorkbook;
export const lexFormula = wasm.lexFormula;
export const lexFormulaPartial = wasm.lexFormulaPartial;
export const parseFormulaPartial = wasm.parseFormulaPartial;
export const rewriteFormulasForCopyDelta = wasm.rewriteFormulasForCopyDelta;
export const canonicalizeFormula = wasm.canonicalizeFormula;
export const localizeFormula = wasm.localizeFormula;

