export type EnsureFormulaWasmNodeBuildOptions = {
  force?: boolean;
};

export type EnsureFormulaWasmNodeBuildResult = {
  outDir: string;
  entryJsPath: string;
  rebuilt: boolean;
};

/**
 * Ensure we have a Node-compatible (`--target nodejs`) wasm-bindgen build of
 * `crates/formula-wasm` available for Vitest/Node consumers.
 */
export function ensureFormulaWasmNodeBuild(options?: EnsureFormulaWasmNodeBuildOptions): EnsureFormulaWasmNodeBuildResult;

/**
 * Returns a `file://` URL to the Node-compatible wasm-bindgen entrypoint
 * (`crates/formula-wasm/pkg-node/formula_wasm.js`).
 */
export function formulaWasmNodeEntryUrl(): string;
