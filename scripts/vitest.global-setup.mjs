import { ensureFormulaWasmNodeBuild } from "./build-formula-wasm-node.mjs";

export default async function globalSetup() {
  // Build once before any vitest suites run so Node can import the wasm-bindgen
  // artifact directly from `crates/formula-wasm/pkg-node/`.
  ensureFormulaWasmNodeBuild();
}

