import { ensureFormulaWasmNodeBuild } from "./build-formula-wasm-node.mjs";

export default async function globalSetup() {
  if (process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true") {
    return;
  }

  ensureFormulaWasmNodeBuild();
}
