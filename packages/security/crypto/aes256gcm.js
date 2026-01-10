import crypto from "node:crypto";
import { assertBufferLength, toBase64, fromBase64 } from "./utils.js";

export const AES_256_GCM = "aes-256-gcm";
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;
export const AES_256_KEY_BYTES = 32;

export function generateAes256Key() {
  return crypto.randomBytes(AES_256_KEY_BYTES);
}

export function encryptAes256Gcm({ plaintext, key, aad = null, iv = null }) {
  if (!Buffer.isBuffer(plaintext)) {
    throw new TypeError("plaintext must be a Buffer");
  }
  assertBufferLength(key, AES_256_KEY_BYTES, "key");
  if (aad !== null && !Buffer.isBuffer(aad)) {
    throw new TypeError("aad must be a Buffer when provided");
  }

  const nonce = iv ?? crypto.randomBytes(AES_GCM_IV_BYTES);
  assertBufferLength(nonce, AES_GCM_IV_BYTES, "iv");

  const cipher = crypto.createCipheriv(AES_256_GCM, key, nonce, {
    authTagLength: AES_GCM_TAG_BYTES
  });

  if (aad) {
    cipher.setAAD(aad);
  }

  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const tag = cipher.getAuthTag();

  return {
    algorithm: AES_256_GCM,
    iv: nonce,
    ciphertext,
    tag
  };
}

export function decryptAes256Gcm({ ciphertext, key, iv, tag, aad = null }) {
  if (!Buffer.isBuffer(ciphertext)) {
    throw new TypeError("ciphertext must be a Buffer");
  }
  assertBufferLength(key, AES_256_KEY_BYTES, "key");
  assertBufferLength(iv, AES_GCM_IV_BYTES, "iv");
  assertBufferLength(tag, AES_GCM_TAG_BYTES, "tag");
  if (aad !== null && !Buffer.isBuffer(aad)) {
    throw new TypeError("aad must be a Buffer when provided");
  }

  const decipher = crypto.createDecipheriv(AES_256_GCM, key, iv, {
    authTagLength: AES_GCM_TAG_BYTES
  });
  decipher.setAuthTag(tag);
  if (aad) {
    decipher.setAAD(aad);
  }

  return Buffer.concat([decipher.update(ciphertext), decipher.final()]);
}

export function serializeEncryptedPayload(payload) {
  return {
    algorithm: payload.algorithm,
    iv: toBase64(payload.iv),
    ciphertext: toBase64(payload.ciphertext),
    tag: toBase64(payload.tag)
  };
}

export function deserializeEncryptedPayload(payload) {
  if (!payload || typeof payload !== "object") {
    throw new TypeError("payload must be an object");
  }

  return {
    algorithm: payload.algorithm,
    iv: fromBase64(payload.iv, "iv"),
    ciphertext: fromBase64(payload.ciphertext, "ciphertext"),
    tag: fromBase64(payload.tag, "tag")
  };
}

