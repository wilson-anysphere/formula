/**
 * @typedef {Object} BinaryStorage
 * @property {() => Promise<Uint8Array | null>} load
 * @property {(data: Uint8Array) => Promise<void>} save
 */

export class InMemoryBinaryStorage {
  constructor() {
    /** @type {Uint8Array | null} */
    this._data = null;
  }

  async load() {
    return this._data ? new Uint8Array(this._data) : null;
  }

  async save(data) {
    this._data = new Uint8Array(data);
  }
}

export class LocalStorageBinaryStorage {
  /**
   * @param {{ workbookId: string, namespace?: string }} opts
   */
  constructor(opts) {
    if (!opts?.workbookId) throw new Error("LocalStorageBinaryStorage requires workbookId");
    const namespace = opts.namespace ?? "formula.ai-rag.sqlite";
    this.key = `${namespace}:${opts.workbookId}`;
  }

  async load() {
    const storage = getLocalStorageOrNull();
    if (!storage) return null;
    const encoded = storage.getItem(this.key);
    if (!encoded) return null;
    return fromBase64(encoded);
  }

  async save(data) {
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    storage.setItem(this.key, toBase64(data));
  }
}

/**
 * localStorage is not always available:
 * - Node >=25 exposes an experimental `globalThis.localStorage` that throws unless
 *   started with `--localstorage-file`.
 * - Vitest's jsdom environment stores the real DOM window on `globalThis.jsdom.window`.
 */
function getLocalStorageOrNull() {
  const isStorage = (value) => value && typeof value.getItem === "function" && typeof value.setItem === "function";

  try {
    // Prefer Vitest's jsdom window when present.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const jsdomStorage = globalThis?.jsdom?.window?.localStorage;
    if (isStorage(jsdomStorage)) return jsdomStorage;
  } catch {
    // ignore
  }

  try {
    // eslint-disable-next-line no-undef
    const windowStorage = typeof window !== "undefined" ? window.localStorage : undefined;
    if (isStorage(windowStorage)) return windowStorage;
  } catch {
    // ignore
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = globalThis?.localStorage;
    return isStorage(storage) ? storage : null;
  } catch {
    return null;
  }
}

/**
 * @param {Uint8Array} data
 */
export function toBase64(data) {
  // Prefer Node's Buffer when available.
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }

  // Browser fallback.
  // Avoid byte-by-byte string concatenation (O(n^2) in many JS engines) by
  // building the binary string in reasonably sized chunks.
  const chunkSize = 0x8000;
  /** @type {string[]} */
  const chunks = [];
  for (let i = 0; i < data.length; i += chunkSize) {
    const chunk = data.subarray(i, i + chunkSize);
    // `String.fromCharCode` expects a list of numbers. Passing a TypedArray via
    // `apply` keeps this browser-safe while avoiding large argument lists.
    // eslint-disable-next-line prefer-spread
    chunks.push(String.fromCharCode.apply(null, chunk));
  }
  // eslint-disable-next-line no-undef
  return btoa(chunks.join(""));
}

/**
 * @param {string} encoded
 */
export function fromBase64(encoded) {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(encoded, "base64"));
  }
  // eslint-disable-next-line no-undef
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  // Fill in chunks to keep the hot loop simple for large payloads.
  const chunkSize = 0x8000;
  for (let i = 0; i < binary.length; i += chunkSize) {
    const end = Math.min(i + chunkSize, binary.length);
    for (let j = i; j < end; j += 1) {
      bytes[j] = binary.charCodeAt(j);
    }
  }
  return bytes;
}
