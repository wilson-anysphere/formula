export function ensureFormulaWasmNodeBuild(options?: { force?: boolean }): {
  outDir: string;
  entryJsPath: string;
  rebuilt: boolean;
};

/**
 * @returns `file://` URL to the JS entry point (`crates/formula-wasm/pkg-node/formula_wasm.js`).
 */
export function formulaWasmNodeEntryUrl(): string;
