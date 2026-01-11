import crypto from "node:crypto";

/**
 * Node.js crypto-backed AES-256-GCM provider for `EncryptedCacheStore`.
 *
 * Hosts still own key management (they provide the key bytes / key version), but
 * this helper avoids re-implementing AES-GCM plumbing in each host.
 *
 * This module imports Node built-ins at evaluation time and should only be used
 * from Node environments (or via the `power-query/src/node.js` entrypoint).
 *
 * @typedef {import("./encryptedStore.js").CacheCryptoProvider} CacheCryptoProvider
 */

const IV_BYTES = 12;
const TAG_BYTES = 16;
const KEY_BYTES = 32;

/**
 * @param {Uint8Array} value
 */
function assertKeyBytes(value) {
  if (!(value instanceof Uint8Array)) {
    throw new TypeError("keyBytes must be a Uint8Array");
  }
  if (value.byteLength !== KEY_BYTES) {
    throw new RangeError(`keyBytes must be ${KEY_BYTES} bytes (got ${value.byteLength})`);
  }
}

/**
 * Create a `CacheCryptoProvider` backed by Node's `crypto` module (AES-256-GCM).
 *
 * @param {{
 *   keyVersion: number;
 *   keyBytes: Uint8Array;
 * }} options
 * @returns {CacheCryptoProvider}
 */
export function createNodeCryptoCacheProvider(options) {
  if (!options || typeof options !== "object") {
    throw new TypeError("options is required");
  }
  if (!Number.isInteger(options.keyVersion) || options.keyVersion < 1) {
    throw new RangeError("keyVersion must be an integer >= 1");
  }
  assertKeyBytes(options.keyBytes);

  const key = Buffer.from(options.keyBytes);

  return {
    keyVersion: options.keyVersion,

    async encryptBytes(plaintext, aad) {
      const iv = crypto.randomBytes(IV_BYTES);
      const cipher = crypto.createCipheriv("aes-256-gcm", key, iv, { authTagLength: TAG_BYTES });
      if (aad) cipher.setAAD(Buffer.from(aad));
      const ciphertext = Buffer.concat([cipher.update(Buffer.from(plaintext)), cipher.final()]);
      const tag = cipher.getAuthTag();
      return {
        keyVersion: options.keyVersion,
        iv: new Uint8Array(iv.buffer, iv.byteOffset, iv.byteLength),
        tag: new Uint8Array(tag.buffer, tag.byteOffset, tag.byteLength),
        ciphertext: new Uint8Array(ciphertext.buffer, ciphertext.byteOffset, ciphertext.byteLength),
      };
    },

    async decryptBytes(payload, aad) {
      const iv = Buffer.from(payload.iv);
      const tag = Buffer.from(payload.tag);
      const ciphertext = Buffer.from(payload.ciphertext);
      const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv, { authTagLength: TAG_BYTES });
      decipher.setAuthTag(tag);
      if (aad) decipher.setAAD(Buffer.from(aad));
      const plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
      return new Uint8Array(plaintext.buffer, plaintext.byteOffset, plaintext.byteLength);
    },
  };
}

