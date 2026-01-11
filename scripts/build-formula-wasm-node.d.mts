export function ensureFormulaWasmNodeBuild(options?: { force?: boolean }): {
  outDir: string;
  entryJsPath: string;
  rebuilt: boolean;
};

/**
 * Returns a `file://` URL to the Node-compatible wasm-bindgen entrypoint
 * (`crates/formula-wasm/pkg-node/formula_wasm.js`).
 */
export function formulaWasmNodeEntryUrl(): string;
