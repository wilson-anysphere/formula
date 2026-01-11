export function defaultWasmModuleUrl(): string {
  return new URL("../pkg/formula_wasm.js", import.meta.url).toString();
}

export function defaultWasmBinaryUrl(): string {
  return new URL("../pkg/formula_wasm_bg.wasm", import.meta.url).toString();
}
