import {
  AES_256_KEY_BYTES,
  decryptAes256Gcm,
  encryptAes256Gcm,
  generateAes256Key,
  serializeEncryptedPayload,
  deserializeEncryptedPayload
} from "../aes256gcm.js";
import { aadFromContext, assertBufferLength, fromBase64, toBase64 } from "../utils.js";

/**
 * Local KMS provider for tests and single-node dev deployments.
 *
 * This implements envelope-style key wrapping:
 * - A Key Encryption Key (KEK) is generated and versioned.
 * - Data Encryption Keys (DEKs) are wrapped with the KEK using AES-256-GCM.
 *
 * In production, replace this provider with AWS/GCP/Azure KMS-backed
 * implementations.
 */
export class LocalKmsProvider {
  constructor({ currentVersion, keksByVersion } = {}) {
    this.provider = "local";
    this.currentVersion = currentVersion ?? 1;
    this.keksByVersion = keksByVersion ?? new Map([[1, generateAes256Key()]]);
    if (!this.keksByVersion.has(this.currentVersion)) {
      throw new Error("keksByVersion must include currentVersion");
    }
    for (const [version, key] of this.keksByVersion.entries()) {
      if (!Number.isInteger(version) || version < 1) {
        throw new RangeError(`Invalid key version: ${String(version)}`);
      }
      assertBufferLength(key, AES_256_KEY_BYTES, `keksByVersion[${version}]`);
    }
  }

  rotateKey() {
    const next = this.currentVersion + 1;
    this.keksByVersion.set(next, generateAes256Key());
    this.currentVersion = next;
    return next;
  }

  _getKek(version) {
    const key = this.keksByVersion.get(version);
    if (!key) {
      throw new Error(`Missing KEK for version ${version}`);
    }
    return key;
  }

  async wrapKey({ plaintextKey, encryptionContext = null }) {
    if (!Buffer.isBuffer(plaintextKey)) {
      throw new TypeError("plaintextKey must be a Buffer");
    }

    const aad = aadFromContext(encryptionContext);
    const encrypted = encryptAes256Gcm({
      plaintext: plaintextKey,
      key: this._getKek(this.currentVersion),
      aad
    });

    return {
      kmsProvider: this.provider,
      kmsKeyVersion: this.currentVersion,
      ...serializeEncryptedPayload(encrypted)
    };
  }

  async unwrapKey({ wrappedKey, encryptionContext = null }) {
    if (!wrappedKey || typeof wrappedKey !== "object") {
      throw new TypeError("wrappedKey must be an object");
    }
    if (wrappedKey.kmsProvider !== this.provider) {
      throw new Error(`Unsupported kmsProvider: ${wrappedKey.kmsProvider}`);
    }

    const aad = aadFromContext(encryptionContext);
    const payload = deserializeEncryptedPayload(wrappedKey);
    return decryptAes256Gcm({
      ...payload,
      key: this._getKek(wrappedKey.kmsKeyVersion),
      aad
    });
  }

  toJSON() {
    const keks = {};
    for (const [version, key] of this.keksByVersion.entries()) {
      keks[String(version)] = toBase64(key);
    }
    return {
      provider: this.provider,
      currentVersion: this.currentVersion,
      keks
    };
  }

  static fromJSON(value) {
    if (!value || typeof value !== "object") {
      throw new TypeError("LocalKmsProvider JSON must be an object");
    }

    const { currentVersion, keks } = value;
    if (!keks || typeof keks !== "object") {
      throw new TypeError("LocalKmsProvider JSON missing keks");
    }

    const keksByVersion = new Map();
    for (const [versionStr, keyBase64] of Object.entries(keks)) {
      const version = Number.parseInt(versionStr, 10);
      if (!Number.isInteger(version) || version < 1) {
        throw new Error(`Invalid key version: ${versionStr}`);
      }
      const decoded = fromBase64(keyBase64, `keks[${versionStr}]`);
      assertBufferLength(decoded, AES_256_KEY_BYTES, `keks[${versionStr}]`);
      keksByVersion.set(version, decoded);
    }

    return new LocalKmsProvider({ currentVersion, keksByVersion });
  }
}
