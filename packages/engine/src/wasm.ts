export function defaultWasmModuleUrl(): string {
  // Vite rewrites `new URL(..., import.meta.url)` into the final built asset URL
  // for both dev and production, while remaining standards-based ESM.
  return new URL("../pkg/formula_wasm.js", import.meta.url).toString();
}

export function defaultWasmBinaryUrl(): string {
  return new URL("../pkg/formula_wasm_bg.wasm", import.meta.url).toString();
}
