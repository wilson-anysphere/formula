/**
 * WebCrypto-backed AES-256-GCM provider for `EncryptedCacheStore`.
 *
 * This helper is intentionally small: hosts still own key management (they provide
 * the key bytes / key version), but don't need to re-implement AES-GCM plumbing.
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
 * @returns {SubtleCrypto}
 */
function requireSubtleCrypto() {
  const subtle = globalThis.crypto?.subtle ?? null;
  if (!subtle) {
    throw new Error("WebCrypto is not available in this environment (crypto.subtle missing)");
  }
  return subtle;
}

/**
 * @returns {(out: Uint8Array) => Uint8Array}
 */
function requireRandomValues() {
  const fn = globalThis.crypto?.getRandomValues?.bind(globalThis.crypto) ?? null;
  if (!fn) {
    throw new Error("WebCrypto is not available in this environment (crypto.getRandomValues missing)");
  }
  return fn;
}

/**
 * Create a `CacheCryptoProvider` backed by WebCrypto AES-256-GCM.
 *
 * @param {{
 *   keyVersion: number;
 *   keyBytes: Uint8Array;
 * }} options
 * @returns {Promise<CacheCryptoProvider>}
 */
export async function createWebCryptoCacheProvider(options) {
  if (!options || typeof options !== "object") {
    throw new TypeError("options is required");
  }
  if (!Number.isInteger(options.keyVersion) || options.keyVersion < 1) {
    throw new RangeError("keyVersion must be an integer >= 1");
  }
  assertKeyBytes(options.keyBytes);

  const subtle = requireSubtleCrypto();
  const getRandomValues = requireRandomValues();
  const key = await subtle.importKey("raw", options.keyBytes, { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);

  return {
    keyVersion: options.keyVersion,

    async encryptBytes(plaintext, aad) {
      const iv = new Uint8Array(IV_BYTES);
      getRandomValues(iv);

      const encrypted = await subtle.encrypt(
        {
          name: "AES-GCM",
          iv,
          additionalData: aad,
          tagLength: TAG_BYTES * 8,
        },
        key,
        plaintext,
      );

      const combined = new Uint8Array(encrypted);
      const ciphertext = combined.subarray(0, Math.max(0, combined.byteLength - TAG_BYTES));
      const tag = combined.subarray(Math.max(0, combined.byteLength - TAG_BYTES));

      return {
        keyVersion: options.keyVersion,
        iv,
        tag,
        ciphertext,
      };
    },

    async decryptBytes(payload, aad) {
      const iv = payload.iv instanceof Uint8Array ? payload.iv : new Uint8Array(payload.iv);
      const tag = payload.tag instanceof Uint8Array ? payload.tag : new Uint8Array(payload.tag);
      const ciphertext = payload.ciphertext instanceof Uint8Array ? payload.ciphertext : new Uint8Array(payload.ciphertext);

      const combined = new Uint8Array(ciphertext.byteLength + tag.byteLength);
      combined.set(ciphertext, 0);
      combined.set(tag, ciphertext.byteLength);

      const plaintext = await subtle.decrypt(
        {
          name: "AES-GCM",
          iv,
          additionalData: aad,
          tagLength: TAG_BYTES * 8,
        },
        key,
        combined,
      );

      return new Uint8Array(plaintext);
    },
  };
}

