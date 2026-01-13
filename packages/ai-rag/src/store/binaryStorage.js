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
  let binary = "";
  for (const byte of data) binary += String.fromCharCode(byte);
  // eslint-disable-next-line no-undef
  return btoa(binary);
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
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
