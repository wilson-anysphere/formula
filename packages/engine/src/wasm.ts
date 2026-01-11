import wasmModuleUrl from "../pkg/formula_wasm.js?url";
import wasmBinaryUrl from "../pkg/formula_wasm_bg.wasm?url";

export function defaultWasmModuleUrl(): string {
  return wasmModuleUrl;
}

export function defaultWasmBinaryUrl(): string {
  return wasmBinaryUrl;
}
