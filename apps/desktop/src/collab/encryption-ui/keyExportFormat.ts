import { AES_256_KEY_BYTES, base64ToBytes, bytesToBase64, utf8Decode, utf8Encode } from "@formula/collab-encryption";

export const ENCRYPTION_KEY_EXPORT_PREFIX = "formula-enc://";
export const ENCRYPTION_KEY_EXPORT_VERSION = 1 as const;

type EncryptionKeyExportPayloadV1 = {
  v: typeof ENCRYPTION_KEY_EXPORT_VERSION;
  docId: string;
  keyId: string;
  keyBytesBase64: string;
};

function bytesToBase64Url(bytes: Uint8Array): string {
  // `bytesToBase64` returns standard base64; normalize to URL-safe base64url without padding.
  return bytesToBase64(bytes).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlToBytes(value: string): Uint8Array {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) throw new Error("Invalid encryption key export string");

  // Convert base64url to base64, restoring padding.
  let base64 = trimmed.replace(/-/g, "+").replace(/_/g, "/");
  const mod = base64.length % 4;
  if (mod === 2) base64 += "==";
  else if (mod === 3) base64 += "=";
  else if (mod !== 0) {
    // `mod===1` is never valid padding.
    throw new Error("Invalid encryption key export string");
  }

  return base64ToBytes(base64);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

export function serializeEncryptionKeyExportString(params: { docId: string; keyId: string; keyBytes: Uint8Array }): string {
  const docId = String(params.docId ?? "").trim();
  const keyId = String(params.keyId ?? "").trim();
  const keyBytes = params.keyBytes;

  if (!docId) throw new Error("docId is required");
  if (!keyId) throw new Error("keyId is required");
  if (!(keyBytes instanceof Uint8Array)) throw new Error("keyBytes must be a Uint8Array");
  if (keyBytes.byteLength !== AES_256_KEY_BYTES) throw new Error("Invalid encryption key length");

  const payload: EncryptionKeyExportPayloadV1 = {
    v: ENCRYPTION_KEY_EXPORT_VERSION,
    docId,
    keyId,
    keyBytesBase64: bytesToBase64(keyBytes),
  };

  const encoded = bytesToBase64Url(utf8Encode(JSON.stringify(payload)));
  return `${ENCRYPTION_KEY_EXPORT_PREFIX}${encoded}`;
}

export function parseEncryptionKeyExportString(raw: string): { docId: string; keyId: string; keyBytes: Uint8Array } {
  const trimmed = String(raw ?? "").trim();
  if (!trimmed) throw new Error("Invalid encryption key export string");

  const token = trimmed.startsWith(ENCRYPTION_KEY_EXPORT_PREFIX)
    ? trimmed.slice(ENCRYPTION_KEY_EXPORT_PREFIX.length).trim()
    : trimmed;

  const bytes = base64UrlToBytes(token);
  let parsed: unknown;
  try {
    parsed = JSON.parse(utf8Decode(bytes));
  } catch {
    throw new Error("Invalid encryption key export string");
  }

  if (!isPlainObject(parsed)) throw new Error("Invalid encryption key export string");
  const v = Number((parsed as any).v);
  if (v !== ENCRYPTION_KEY_EXPORT_VERSION) throw new Error("Unsupported encryption key export version");

  const docId = String((parsed as any).docId ?? "").trim();
  const keyId = String((parsed as any).keyId ?? "").trim();
  const keyBytesBase64 = String((parsed as any).keyBytesBase64 ?? "").trim();
  if (!docId || !keyId || !keyBytesBase64) throw new Error("Invalid encryption key export string");

  let keyBytes: Uint8Array;
  try {
    keyBytes = base64ToBytes(keyBytesBase64);
  } catch {
    throw new Error("Invalid encryption key export string");
  }
  if (keyBytes.byteLength !== AES_256_KEY_BYTES) throw new Error("Invalid encryption key length");

  return { docId, keyId, keyBytes };
}

