export const CELL_ENCRYPTION_VERSION = 1 as const;
export const CELL_ENCRYPTION_ALG = "AES-256-GCM" as const;
export const AES_256_KEY_BYTES = 32;
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;

export interface EncryptedCellPayloadV1 {
  v: typeof CELL_ENCRYPTION_VERSION;
  alg: typeof CELL_ENCRYPTION_ALG;
  keyId: string;
  ivBase64: string;
  tagBase64: string;
  ciphertextBase64: string;
}

export type EncryptedCellPayload = EncryptedCellPayloadV1;

export interface CellPlaintext {
  value: unknown;
  formula: string | null;
  // Optional per-cell formatting payload. `@formula/collab-session` and the desktop
  // binder can opt into encrypting this field via `encryption.encryptFormat`.
  format?: unknown;
}

export interface CellEncryptionContext {
  docId: string;
  sheetId: string;
  row: number;
  col: number;
}

export interface CellEncryptionKey {
  keyId: string;
  keyBytes: Uint8Array;
}

export function isEncryptedCellPayload(value: unknown): value is EncryptedCellPayloadV1 {
  if (!value || typeof value !== "object") return false;
  const v = value as any;
  return (
    v.v === CELL_ENCRYPTION_VERSION &&
    v.alg === CELL_ENCRYPTION_ALG &&
    typeof v.keyId === "string" &&
    typeof v.ivBase64 === "string" &&
    typeof v.tagBase64 === "string" &&
    typeof v.ciphertextBase64 === "string"
  );
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && (value as any).constructor === Object;
}

function sortJson(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(sortJson);
  }
  if (isPlainObject(value)) {
    const sorted: Record<string, unknown> = {};
    for (const key of Object.keys(value).sort()) {
      sorted[key] = sortJson(value[key]);
    }
    return sorted;
  }
  return value;
}

/**
 * Deterministic JSON encoding suitable for use as AAD / encryption context.
 *
 * This intentionally does *not* attempt to be a general-purpose canonicalization
 * algorithm. It exists so that `{ docId, sheetId, row, col }` produces identical
 * bytes across runtimes (Node/browser) for AES-GCM additional authenticated data.
 */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(sortJson(value));
}

export function aadBytesFromContext(context: CellEncryptionContext): Uint8Array {
  return utf8Encode(canonicalJson(context));
}

export function assertKeyBytes(keyBytes: Uint8Array): void {
  if (!(keyBytes instanceof Uint8Array)) {
    throw new TypeError("keyBytes must be a Uint8Array");
  }
  if (keyBytes.byteLength !== AES_256_KEY_BYTES) {
    throw new RangeError(`keyBytes must be ${AES_256_KEY_BYTES} bytes (got ${keyBytes.byteLength})`);
  }
}

export function assertIvBytes(iv: Uint8Array): void {
  if (!(iv instanceof Uint8Array)) {
    throw new TypeError("iv must be a Uint8Array");
  }
  if (iv.byteLength !== AES_GCM_IV_BYTES) {
    throw new RangeError(`iv must be ${AES_GCM_IV_BYTES} bytes (got ${iv.byteLength})`);
  }
}

export function assertTagBytes(tag: Uint8Array): void {
  if (!(tag instanceof Uint8Array)) {
    throw new TypeError("tag must be a Uint8Array");
  }
  if (tag.byteLength !== AES_GCM_TAG_BYTES) {
    throw new RangeError(`tag must be ${AES_GCM_TAG_BYTES} bytes (got ${tag.byteLength})`);
  }
}

export function concatBytes(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.byteLength + b.byteLength);
  out.set(a, 0);
  out.set(b, a.byteLength);
  return out;
}

export function bytesToBase64(bytes: Uint8Array): string {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError("bytesToBase64 expects a Uint8Array");
  }
  // Node.js: Buffer is the most efficient.
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  // Browsers: btoa works on binary strings.
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  // eslint-disable-next-line no-undef
  return btoa(bin);
}

export function base64ToBytes(value: string): Uint8Array {
  if (typeof value !== "string") {
    throw new TypeError("base64ToBytes expects a base64 string");
  }
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(value, "base64"));
  }

  // eslint-disable-next-line no-undef
  const bin = atob(value);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
  return out;
}

export function utf8Encode(text: string): Uint8Array {
  const encoder = typeof TextEncoder !== "undefined" ? new TextEncoder() : null;
  if (encoder) return encoder.encode(text);
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(text, "utf8"));
  throw new Error("No UTF-8 encoder available (TextEncoder missing)");
}

export function utf8Decode(bytes: Uint8Array): string {
  const decoder = typeof TextDecoder !== "undefined" ? new TextDecoder() : null;
  if (decoder) return decoder.decode(bytes);
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("utf8");
  throw new Error("No UTF-8 decoder available (TextDecoder missing)");
}
