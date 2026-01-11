import createHarfBuzzModule from "harfbuzzjs/hb.js";
import hbjs from "harfbuzzjs/hbjs.js";

/**
 * @typedef {ReturnType<typeof hbjs>} HarfBuzz
 */

const HB_WASM_URL = new URL("./hb.wasm", import.meta.url);

/** @type {Promise<HarfBuzz> | null} */
let HARFBUZZ_PROMISE = null;

/**
 * @param {URL} url
 * @returns {Promise<ArrayBuffer>}
 */
async function readUrlAsArrayBuffer(url) {
  // Node's `fetch()` doesn't support `file:` URLs (and we want this module to work in both Node and browsers).
  if (url.protocol === "file:") {
    // Use dynamic specifiers to keep browser bundlers from trying to resolve Node builtins.
    const { readFile } = await import("node:" + "fs/promises");
    const { fileURLToPath } = await import("node:" + "url");
    const buf = await readFile(fileURLToPath(url));
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  }

  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch HarfBuzz WASM (${res.status} ${res.statusText})`);
  return await res.arrayBuffer();
}

/**
 * Load the HarfBuzz WASM module (cached singleton).
 *
 * Consumers should call this once at startup and reuse the returned instance.
 *
 * @returns {Promise<HarfBuzz>}
 */
export function loadHarfBuzz() {
  if (!HARFBUZZ_PROMISE) {
    HARFBUZZ_PROMISE = (async () => {
      const wasmBinary = await readUrlAsArrayBuffer(HB_WASM_URL);
      const moduleInstance = await createHarfBuzzModule({ wasmBinary });
      return hbjs(moduleInstance);
    })();
  }
  return HARFBUZZ_PROMISE;
}
