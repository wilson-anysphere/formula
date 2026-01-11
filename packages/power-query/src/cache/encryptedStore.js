/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 * @typedef {import("./cache.js").CacheStore} CacheStore
 */

/**
 * Host-provided crypto provider for encrypting cached values at rest.
 *
 * Power Query cache values can include Arrow IPC payloads (`Uint8Array`). To avoid
 * base64 overhead, encryption is performed on bytes and ciphertext is stored as a
 * `Uint8Array`.
 *
 * Implementations must use AES-256-GCM and authenticate the provided AAD.
 *
 * @typedef {{
 *   keyVersion: number;
 *   encryptBytes: (plaintext: Uint8Array, aad?: Uint8Array) => Promise<{
 *     keyVersion: number;
 *     iv: Uint8Array;
 *     tag: Uint8Array;
 *     ciphertext: Uint8Array;
 *   }>;
 *   decryptBytes: (
 *     payload: { keyVersion: number; iv: Uint8Array; tag: Uint8Array; ciphertext: Uint8Array },
 *     aad?: Uint8Array,
 *   ) => Promise<Uint8Array>;
 * }} CacheCryptoProvider
 */

const AAD_SCOPE = "power-query-cache";
const ENVELOPE_MARKER = "power-query-cache-encrypted";
const ENVELOPE_VERSION = 1;

const VALUE_MAGIC = /** @type {const} */ ([0x50, 0x51, 0x43, 0x56]); // "PQCV"
const VALUE_VERSION = 1;
const TYPE_KEY = "__pq_cache_type";

function isPlainObject(value) {
  return value != null && typeof value === "object" && /** @type {any} */ (value).constructor === Object;
}

function sortJson(value) {
  if (Array.isArray(value)) {
    return value.map(sortJson);
  }
  if (isPlainObject(value)) {
    const sorted = {};
    for (const key of Object.keys(value).sort()) {
      // @ts-ignore - runtime
      sorted[key] = sortJson(value[key]);
    }
    return sorted;
  }
  return value;
}

/**
 * Deterministic JSON encoding suitable for AAD / encryption context.
 *
 * Do NOT use this for security-sensitive canonicalization of untrusted input; it
 * exists so encryption context bytes are stable across runtime instances.
 *
 * @param {unknown} value
 */
function canonicalJson(value) {
  return JSON.stringify(sortJson(value));
}

function utf8Encode(text) {
  if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(text);
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(text, "utf8"));
  throw new Error("No UTF-8 encoder available (TextEncoder missing)");
}

function utf8Decode(bytes) {
  if (typeof TextDecoder !== "undefined") return new TextDecoder().decode(bytes);
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("utf8");
  throw new Error("No UTF-8 decoder available (TextDecoder missing)");
}

/**
 * Normalize a byte buffer into a plain `Uint8Array` view (not a Node Buffer).
 *
 * This avoids `Buffer#toJSON()` behavior when persisting encrypted payloads via
 * JSON (e.g. `FileSystemCacheStore`).
 *
 * @param {Uint8Array | ArrayBuffer} value
 * @returns {Uint8Array}
 */
function normalizeBytes(value) {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  return new Uint8Array(value);
}

/**
 * @param {Uint8Array} out
 * @param {number} offset
 * @param {number} value
 */
function writeUint32LE(out, offset, value) {
  const view = new DataView(out.buffer, out.byteOffset, out.byteLength);
  view.setUint32(offset, value >>> 0, true);
}

/**
 * @param {DataView} view
 * @param {number} offset
 */
function readUint32LE(view, offset) {
  return view.getUint32(offset, true);
}

/**
 * Encode a structured-cloneable value into bytes without base64 encoding nested
 * `Uint8Array` values.
 *
 * The format is intentionally simple and versioned:
 *
 *   [4 bytes magic "PQCV"][1 byte version]
 *   [u32 jsonLength][json utf8 bytes]
 *   [u32 binCount][repeated: u32 binLength + bin bytes]
 *
 * JSON is used for the general shape and scalar values, while binary blobs are
 * extracted into a separate section referenced from JSON via tagged objects.
 *
 * @param {unknown} value
 * @returns {Uint8Array}
 */
function encodeValueBytes(value) {
  /** @type {Uint8Array[]} */
  const bins = [];
  const encodedForJson = encodeForJson(value, bins, new WeakSet());
  const jsonBytes = utf8Encode(JSON.stringify(encodedForJson));

  let totalLen = 4 + 1 + 4 + jsonBytes.byteLength + 4;
  for (const bin of bins) totalLen += 4 + bin.byteLength;

  const out = new Uint8Array(totalLen);
  out.set(VALUE_MAGIC, 0);
  out[4] = VALUE_VERSION;

  let offset = 5;
  writeUint32LE(out, offset, jsonBytes.byteLength);
  offset += 4;
  out.set(jsonBytes, offset);
  offset += jsonBytes.byteLength;

  writeUint32LE(out, offset, bins.length);
  offset += 4;

  for (const bin of bins) {
    writeUint32LE(out, offset, bin.byteLength);
    offset += 4;
    out.set(bin, offset);
    offset += bin.byteLength;
  }

  return out;
}

/**
 * @param {Uint8Array} bytes
 * @returns {unknown}
 */
function decodeValueBytes(bytes) {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError("Expected Uint8Array plaintext bytes");
  }

  if (bytes.byteLength < 5) {
    throw new Error("Invalid cache plaintext (truncated)");
  }

  for (let i = 0; i < VALUE_MAGIC.length; i++) {
    if (bytes[i] !== VALUE_MAGIC[i]) {
      throw new Error("Invalid cache plaintext (bad magic)");
    }
  }

  const version = bytes[4];
  if (version !== VALUE_VERSION) {
    throw new Error(`Unsupported cache plaintext version '${version}'`);
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 5;

  if (offset + 4 > bytes.byteLength) {
    throw new Error("Invalid cache plaintext (missing JSON length)");
  }
  const jsonLen = readUint32LE(view, offset);
  offset += 4;

  if (offset + jsonLen > bytes.byteLength) {
    throw new Error("Invalid cache plaintext (JSON out of bounds)");
  }

  const jsonText = utf8Decode(bytes.subarray(offset, offset + jsonLen));
  offset += jsonLen;

  const parsed = JSON.parse(jsonText);

  if (offset + 4 > bytes.byteLength) {
    throw new Error("Invalid cache plaintext (missing binary count)");
  }
  const binCount = readUint32LE(view, offset);
  offset += 4;

  /** @type {Uint8Array[]} */
  const bins = [];
  for (let i = 0; i < binCount; i++) {
    if (offset + 4 > bytes.byteLength) throw new Error("Invalid cache plaintext (missing binary length)");
    const binLen = readUint32LE(view, offset);
    offset += 4;
    if (offset + binLen > bytes.byteLength) throw new Error("Invalid cache plaintext (binary out of bounds)");
    bins.push(bytes.subarray(offset, offset + binLen));
    offset += binLen;
  }

  return decodeFromJson(parsed, bins);
}

/**
 * @param {unknown} value
 * @param {Uint8Array[]} bins
 * @param {WeakSet<object>} seen
 * @returns {unknown}
 */
function encodeForJson(value, bins, seen) {
  if (value === null) return null;

  const type = typeof value;
  if (type === "string" || type === "boolean") return value;

  if (type === "number") {
    if (Number.isFinite(value)) return value;
    return { [TYPE_KEY]: "number", value: String(value) };
  }

  if (type === "bigint") {
    return { [TYPE_KEY]: "bigint", value: value.toString() };
  }

  if (type === "undefined") {
    return { [TYPE_KEY]: "undefined" };
  }

  if (value instanceof Date) {
    if (!Number.isNaN(value.getTime())) return { [TYPE_KEY]: "date", value: value.toISOString() };
    return { [TYPE_KEY]: "date", value: null };
  }

  if (value instanceof Uint8Array) {
    const idx = bins.push(value) - 1;
    return { [TYPE_KEY]: "u8", i: idx };
  }

  if (value instanceof ArrayBuffer) {
    const idx = bins.push(new Uint8Array(value)) - 1;
    return { [TYPE_KEY]: "u8", i: idx };
  }

  if (Array.isArray(value)) {
    return value.map((item) => encodeForJson(item, bins, seen));
  }

  if (value instanceof Map) {
    return {
      [TYPE_KEY]: "map",
      entries: Array.from(value.entries()).map(([k, v]) => [encodeForJson(k, bins, seen), encodeForJson(v, bins, seen)]),
    };
  }

  if (value instanceof Set) {
    return {
      [TYPE_KEY]: "set",
      entries: Array.from(value.values()).map((v) => encodeForJson(v, bins, seen)),
    };
  }

  if (type === "object") {
    const obj = /** @type {object} */ (value);
    if (seen.has(obj)) {
      throw new Error("Cannot cache values with circular references");
    }
    seen.add(obj);
    const out = {};
    const keys = Object.keys(obj).sort();
    for (const key of keys) {
      // @ts-ignore - runtime indexing
      out[key] = encodeForJson(obj[key], bins, seen);
    }
    seen.delete(obj);
    return out;
  }

  throw new Error(`Unsupported cache value type '${type}'`);
}

/**
 * @param {unknown} value
 * @param {Uint8Array[]} bins
 * @returns {unknown}
 */
function decodeFromJson(value, bins) {
  if (Array.isArray(value)) {
    return value.map((item) => decodeFromJson(item, bins));
  }

  if (value && typeof value === "object") {
    if (!Array.isArray(value) && TYPE_KEY in value) {
      // @ts-ignore - runtime
      const type = value[TYPE_KEY];

      if (type === "u8" && typeof value.i === "number") {
        const bin = bins[value.i] ?? null;
        if (!bin) throw new Error("Invalid cache plaintext (missing binary segment)");
        return bin;
      }

      if (type === "date") {
        if (typeof value.value === "string") {
          const parsed = new Date(value.value);
          if (!Number.isNaN(parsed.getTime())) return parsed;
        }
        return null;
      }

      if (type === "bigint" && typeof value.value === "string") {
        return BigInt(value.value);
      }

      if (type === "undefined") {
        return undefined;
      }

      if (type === "number" && typeof value.value === "string") {
        switch (value.value) {
          case "NaN":
            return Number.NaN;
          case "Infinity":
            return Number.POSITIVE_INFINITY;
          case "-Infinity":
            return Number.NEGATIVE_INFINITY;
          default: {
            const parsed = Number(value.value);
            return Number.isNaN(parsed) ? null : parsed;
          }
        }
      }

      if (type === "map" && Array.isArray(value.entries)) {
        const out = new Map();
        for (const entry of value.entries) {
          if (!Array.isArray(entry) || entry.length !== 2) continue;
          out.set(decodeFromJson(entry[0], bins), decodeFromJson(entry[1], bins));
        }
        return out;
      }

      if (type === "set" && Array.isArray(value.entries)) {
        const out = new Set();
        for (const entry of value.entries) {
          out.add(decodeFromJson(entry, bins));
        }
        return out;
      }
    }

    for (const [k, v] of Object.entries(value)) {
      // @ts-ignore - runtime indexing
      value[k] = decodeFromJson(v, bins);
    }
  }

  return value;
}

/**
 * @typedef {{
 *   __pq_cache_encrypted: typeof ENVELOPE_MARKER;
 *   v: number;
  *   payload: { keyVersion: number; iv: Uint8Array | ArrayBuffer; tag: Uint8Array | ArrayBuffer; ciphertext: Uint8Array | ArrayBuffer };
 * }} EncryptedCacheEnvelopeV1
 */

/**
 * @param {unknown} value
 * @returns {number | null}
 */
function readEnvelopeVersion(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  // @ts-ignore - runtime access
  if (value.__pq_cache_encrypted !== ENVELOPE_MARKER) return null;
  // @ts-ignore - runtime access
  const v = value.v;
  return Number.isInteger(v) && v >= 1 ? v : null;
}

/**
 * @param {unknown} value
 * @returns {value is EncryptedCacheEnvelopeV1}
 */
function isEncryptedEnvelopeV1(value) {
  if (readEnvelopeVersion(value) !== ENVELOPE_VERSION) return false;
  // @ts-ignore - runtime access
  const payload = value.payload;
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) return false;
  // @ts-ignore - runtime access
  if (typeof payload.keyVersion !== "number") return false;
  // @ts-ignore - runtime access
  const { iv, tag, ciphertext } = payload;
  const isBytes = (v) => v instanceof Uint8Array || v instanceof ArrayBuffer;
  return isBytes(iv) && isBytes(tag) && isBytes(ciphertext);
}

/**
 * @param {string | undefined} storeId
 * @param {number} schemaVersion
 * @returns {Uint8Array}
 */
function buildAadBytes(storeId, schemaVersion) {
  /** @type {any} */
  const aad = { scope: AAD_SCOPE, schemaVersion };
  if (storeId != null) aad.storeId = storeId;
  return utf8Encode(canonicalJson(aad));
}

export class EncryptedCacheStore {
  /**
   * @param {{
   *   store: CacheStore;
   *   crypto: CacheCryptoProvider;
   *   storeId?: string;
   * }} options
   */
  constructor(options) {
    this.store = options.store;
    this.crypto = options.crypto;
    this.storeId = options.storeId;
    /** @type {Map<number, Uint8Array>} */
    this._aadBytes = new Map();
  }

  /**
   * @param {number} schemaVersion
   */
  aad(schemaVersion) {
    const existing = this._aadBytes.get(schemaVersion);
    if (existing) return existing;
    const bytes = buildAadBytes(this.storeId, schemaVersion);
    this._aadBytes.set(schemaVersion, bytes);
    return bytes;
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    const entry = await this.store.get(key);
    if (!entry) return null;

    const envelopeVersion = readEnvelopeVersion(entry.value);
    if (envelopeVersion == null) {
      // Backwards-compat: tolerate plaintext entries without crashing.
      // Best-effort delete so plaintext doesn't linger once encryption is enabled.
      await this.store.delete(key).catch(() => {});
      return null;
    }

    // Forward-compat: keep unknown encrypted envelope versions so downgrades do
    // not delete data that newer versions might still understand.
    if (envelopeVersion !== ENVELOPE_VERSION) {
      return null;
    }

    if (!isEncryptedEnvelopeV1(entry.value)) {
      await this.store.delete(key).catch(() => {});
      return null;
    }

    const envelope = /** @type {EncryptedCacheEnvelopeV1} */ (entry.value);

    let plaintext;
    try {
      const payload = envelope.payload;
      plaintext = await this.crypto.decryptBytes(
        {
          keyVersion: payload.keyVersion,
          iv: payload.iv instanceof Uint8Array ? payload.iv : new Uint8Array(payload.iv),
          tag: payload.tag instanceof Uint8Array ? payload.tag : new Uint8Array(payload.tag),
          ciphertext: payload.ciphertext instanceof Uint8Array ? payload.ciphertext : new Uint8Array(payload.ciphertext),
        },
        this.aad(envelope.v),
      );
    } catch {
      await this.store.delete(key).catch(() => {});
      return null;
    }

    let value;
    try {
      value = decodeValueBytes(plaintext);
    } catch {
      await this.store.delete(key).catch(() => {});
      return null;
    }

    return { ...entry, value };
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    const plaintext = encodeValueBytes(entry.value);
    const aad = this.aad(ENVELOPE_VERSION);

    const payload = await this.crypto.encryptBytes(plaintext, aad);
    const normalizedPayload = {
      keyVersion: payload.keyVersion,
      iv: normalizeBytes(payload.iv),
      tag: normalizeBytes(payload.tag),
      ciphertext: normalizeBytes(payload.ciphertext),
    };

    const envelope = {
      __pq_cache_encrypted: ENVELOPE_MARKER,
      v: ENVELOPE_VERSION,
      payload: normalizedPayload,
    };

    await this.store.set(key, { ...entry, value: envelope });
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    await this.store.delete(key);
  }

  async clear() {
    if (this.store.clear) await this.store.clear();
  }

  /**
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs) {
    if (this.store.pruneExpired) await this.store.pruneExpired(nowMs);
  }
}
