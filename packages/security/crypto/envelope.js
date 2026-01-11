import { decryptAes256Gcm, encryptAes256Gcm, generateAes256Key, serializeEncryptedPayload, deserializeEncryptedPayload } from "./aes256gcm.js";
import { aadFromContext } from "./utils.js";

/**
 * Encrypt data using a per-object DEK (data encryption key), wrapped via the
 * supplied KMS provider (envelope encryption).
 */
export async function encryptEnvelope({ plaintext, kmsProvider, encryptionContext = null }) {
  if (!kmsProvider || typeof kmsProvider.wrapKey !== "function") {
    throw new TypeError("kmsProvider must implement wrapKey()");
  }
  if (!Buffer.isBuffer(plaintext)) {
    throw new TypeError("plaintext must be a Buffer");
  }

  const dek = generateAes256Key();
  const aad = aadFromContext(encryptionContext);

  const encrypted = encryptAes256Gcm({
    plaintext,
    key: dek,
    aad
  });

  const wrappedDek = await kmsProvider.wrapKey({
    plaintextKey: dek,
    encryptionContext
  });

  return {
    schemaVersion: 1,
    wrappedDek,
    ...serializeEncryptedPayload(encrypted)
  };
}

export async function decryptEnvelope({ encryptedEnvelope, kmsProvider, encryptionContext = null }) {
  if (!kmsProvider || typeof kmsProvider.unwrapKey !== "function") {
    throw new TypeError("kmsProvider must implement unwrapKey()");
  }
  if (!encryptedEnvelope || typeof encryptedEnvelope !== "object") {
    throw new TypeError("encryptedEnvelope must be an object");
  }

  const dek = await kmsProvider.unwrapKey({
    wrappedKey: encryptedEnvelope.wrappedDek,
    encryptionContext
  });

  const aad = aadFromContext(encryptionContext);
  const payload = deserializeEncryptedPayload(encryptedEnvelope);

  return decryptAes256Gcm({
    ...payload,
    key: dek,
    aad
  });
}
