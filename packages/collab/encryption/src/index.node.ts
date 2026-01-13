import crypto from "node:crypto";

// The Node entrypoint uses `node:crypto` directly (not WebCrypto), but the
// desktop binder code path uses the WebCrypto implementation in `index.node.js`,
// which caches imported `CryptoKey`s. Re-export a clear helper so callers can
// release key material on teardown without knowing which implementation is active.
import { clearEncryptionKeyCache as clearWebCryptoKeyCache } from "./index.node.js";

import {
  AES_GCM_IV_BYTES,
  AES_GCM_TAG_BYTES,
  CELL_ENCRYPTION_ALG,
  CELL_ENCRYPTION_VERSION,
  aadBytesFromContext,
  assertKeyBytes,
  base64ToBytes,
  bytesToBase64,
  type CellEncryptionContext,
  type CellEncryptionKey,
  type CellPlaintext,
  type EncryptedCellPayloadV1,
} from "./shared.ts";

export async function encryptCellPlaintext(opts: {
  plaintext: CellPlaintext;
  key: CellEncryptionKey;
  context: CellEncryptionContext;
}): Promise<EncryptedCellPayloadV1> {
  const { plaintext, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  const iv = crypto.randomBytes(AES_GCM_IV_BYTES);
  const aad = aadBytesFromContext(context);

  const json = JSON.stringify(plaintext);
  const bytes = Buffer.from(json, "utf8");

  const cipher = crypto.createCipheriv("aes-256-gcm", Buffer.from(key.keyBytes), iv, {
    authTagLength: AES_GCM_TAG_BYTES,
  });
  cipher.setAAD(Buffer.from(aad));
  const ciphertext = Buffer.concat([cipher.update(bytes), cipher.final()]);
  const tag = cipher.getAuthTag();

  return {
    v: CELL_ENCRYPTION_VERSION,
    alg: CELL_ENCRYPTION_ALG,
    keyId: key.keyId,
    ivBase64: bytesToBase64(new Uint8Array(iv)),
    tagBase64: bytesToBase64(new Uint8Array(tag)),
    ciphertextBase64: bytesToBase64(new Uint8Array(ciphertext)),
  };
}

export async function decryptCellPlaintext(opts: {
  encrypted: EncryptedCellPayloadV1;
  key: CellEncryptionKey;
  context: CellEncryptionContext;
}): Promise<CellPlaintext> {
  const { encrypted, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  if (encrypted.keyId !== key.keyId) {
    throw new Error(`Key id mismatch (payload=${encrypted.keyId}, resolver=${key.keyId})`);
  }

  const iv = base64ToBytes(encrypted.ivBase64);
  const tag = base64ToBytes(encrypted.tagBase64);
  const ciphertext = base64ToBytes(encrypted.ciphertextBase64);

  const aad = aadBytesFromContext(context);

  const decipher = crypto.createDecipheriv("aes-256-gcm", Buffer.from(key.keyBytes), Buffer.from(iv), {
    authTagLength: AES_GCM_TAG_BYTES,
  });
  decipher.setAuthTag(Buffer.from(tag));
  decipher.setAAD(Buffer.from(aad));
  const plaintextBytes = Buffer.concat([
    decipher.update(Buffer.from(ciphertext)),
    decipher.final(),
  ]);

  const json = plaintextBytes.toString("utf8");
  return JSON.parse(json) as CellPlaintext;
}

export function clearEncryptionKeyCache(): void {
  clearWebCryptoKeyCache();
}

export {
  AES_256_KEY_BYTES,
  AES_GCM_IV_BYTES,
  AES_GCM_TAG_BYTES,
  CELL_ENCRYPTION_ALG,
  CELL_ENCRYPTION_VERSION,
  aadBytesFromContext,
  base64ToBytes,
  bytesToBase64,
  concatBytes,
  canonicalJson,
  isEncryptedCellPayload,
  utf8Decode,
  utf8Encode,
} from "./shared.ts";
export type {
  CellEncryptionContext,
  CellEncryptionKey,
  CellPlaintext,
  EncryptedCellPayload,
  EncryptedCellPayloadV1,
} from "./shared.ts";
