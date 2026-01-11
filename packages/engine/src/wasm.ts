const DEFAULT_WASM_WRAPPER_PATH = "/engine/formula_wasm.js";
const DEFAULT_WASM_BINARY_PATH = "/engine/formula_wasm_bg.wasm";

function resolvePublicUrl(assetPath: string): string {
  const location = (globalThis as unknown as { location?: Location }).location;
  if (!location) {
    throw new Error("defaultWasm*Url() requires a browser/worker runtime with `location`.");
  }

  return new URL(assetPath, location.origin).toString();
}

/**
 * URL to the wasm-bindgen JS wrapper in app `public/engine/*`.
 *
 * We keep this path stable so the worker can `import()` it at runtime in both
 * dev and production builds.
 */
export function defaultWasmModuleUrl(): string {
  return resolvePublicUrl(DEFAULT_WASM_WRAPPER_PATH);
}

/**
 * URL to the `.wasm` binary in app `public/engine/*`.
 *
 * The worker passes this to the wasm-bindgen init function so builds remain
 * robust even if the wrapper can't derive the correct URL from `import.meta.url`
 * (e.g. when assets are fingerprinted).
 */
export function defaultWasmBinaryUrl(): string {
  return resolvePublicUrl(DEFAULT_WASM_BINARY_PATH);
}
