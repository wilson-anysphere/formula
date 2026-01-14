// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`.
//
// This is loaded by `packages/engine/src/engine.worker.ts` via dynamic import (runtime string),
// so it must be plain JS (not TS) to avoid relying on Vite/Vitest transforms.

export default async function init() {
  // No-op.
}

export function supportedLocaleIds() {
  // Deterministic ordering (the real WASM export sorts these).
  return ["de-DE", "en-US", "es-ES", "fr-FR", "ja-JP"];
}

export function getLocaleInfo(localeId) {
  if (localeId !== "de-DE") {
    throw new Error(`unknown localeId: ${localeId}`);
  }

  return {
    localeId: "de-DE",
    decimalSeparator: ",",
    argSeparator: ";",
    arrayRowSeparator: ";",
    arrayColSeparator: "\\",
    thousandsSeparator: ".",
    isRtl: false,
    booleanTrue: "WAHR",
    booleanFalse: "FALSCH"
  };
}

