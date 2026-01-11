import crypto from "node:crypto";
import { assertBufferLength } from "./utils";

export const AES_256_GCM = "aes-256-gcm";
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;
export const AES_256_KEY_BYTES = 32;

export function generateAes256Key(): Buffer {
  return crypto.randomBytes(AES_256_KEY_BYTES);
}

export function encryptAes256Gcm({
  plaintext,
  key,
  aad = null,
  iv = null
}: {
  plaintext: Buffer;
  key: Buffer;
  aad?: Buffer | null;
  iv?: Buffer | null;
}): {
  algorithm: typeof AES_256_GCM;
  iv: Buffer;
  ciphertext: Buffer;
  tag: Buffer;
} {
  if (!Buffer.isBuffer(plaintext)) {
    throw new TypeError("plaintext must be a Buffer");
  }
  assertBufferLength(key, AES_256_KEY_BYTES, "key");
  if (aad !== null && !Buffer.isBuffer(aad)) {
    throw new TypeError("aad must be a Buffer when provided");
  }

  const nonce = iv ?? crypto.randomBytes(AES_GCM_IV_BYTES);
  assertBufferLength(nonce, AES_GCM_IV_BYTES, "iv");

  const cipher = crypto.createCipheriv(AES_256_GCM, key, nonce, { authTagLength: AES_GCM_TAG_BYTES });
  if (aad) {
    cipher.setAAD(aad);
  }

  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const tag = cipher.getAuthTag();

  return { algorithm: AES_256_GCM, iv: nonce, ciphertext, tag };
}

export function decryptAes256Gcm({
  ciphertext,
  key,
  iv,
  tag,
  aad = null
}: {
  ciphertext: Buffer;
  key: Buffer;
  iv: Buffer;
  tag: Buffer;
  aad?: Buffer | null;
}): Buffer {
  if (!Buffer.isBuffer(ciphertext)) {
    throw new TypeError("ciphertext must be a Buffer");
  }
  assertBufferLength(key, AES_256_KEY_BYTES, "key");
  assertBufferLength(iv, AES_GCM_IV_BYTES, "iv");
  assertBufferLength(tag, AES_GCM_TAG_BYTES, "tag");
  if (aad !== null && !Buffer.isBuffer(aad)) {
    throw new TypeError("aad must be a Buffer when provided");
  }

  const decipher = crypto.createDecipheriv(AES_256_GCM, key, iv, { authTagLength: AES_GCM_TAG_BYTES });
  decipher.setAuthTag(tag);
  if (aad) {
    decipher.setAAD(aad);
  }

  return Buffer.concat([decipher.update(ciphertext), decipher.final()]);
}

