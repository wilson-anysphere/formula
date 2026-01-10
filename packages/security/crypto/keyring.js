import {
  AES_256_KEY_BYTES,
  decryptAes256Gcm,
  encryptAes256Gcm,
  generateAes256Key,
  serializeEncryptedPayload,
  deserializeEncryptedPayload
} from "./aes256gcm.js";
import { aadFromContext, assertBufferLength, fromBase64, toBase64 } from "./utils.js";

export class KeyRing {
  constructor({ currentVersion, keysByVersion }) {
    if (!Number.isInteger(currentVersion) || currentVersion < 1) {
      throw new RangeError("currentVersion must be an integer >= 1");
    }
    if (!(keysByVersion instanceof Map)) {
      throw new TypeError("keysByVersion must be a Map");
    }
    if (!keysByVersion.has(currentVersion)) {
      throw new Error("keysByVersion must include currentVersion");
    }
    for (const [version, key] of keysByVersion.entries()) {
      if (!Number.isInteger(version) || version < 1) {
        throw new RangeError(`Invalid key version: ${String(version)}`);
      }
      assertBufferLength(key, AES_256_KEY_BYTES, `keysByVersion[${version}]`);
    }
    this.currentVersion = currentVersion;
    this.keysByVersion = keysByVersion;
  }

  static create() {
    const key = generateAes256Key();
    return new KeyRing({
      currentVersion: 1,
      keysByVersion: new Map([[1, key]])
    });
  }

  rotate() {
    const nextVersion = this.currentVersion + 1;
    this.keysByVersion.set(nextVersion, generateAes256Key());
    this.currentVersion = nextVersion;
    return nextVersion;
  }

  getKey(version) {
    const key = this.keysByVersion.get(version);
    if (!key) {
      throw new Error(`Missing key material for version ${version}`);
    }
    return key;
  }

  /**
   * Encrypt bytes with the current key version.
   */
  encrypt(plaintext, { aadContext = null } = {}) {
    const aad = aadFromContext(aadContext);
    const encrypted = encryptAes256Gcm({
      plaintext,
      key: this.getKey(this.currentVersion),
      aad
    });

    return {
      keyVersion: this.currentVersion,
      ...serializeEncryptedPayload(encrypted)
    };
  }

  decrypt(encrypted, { aadContext = null } = {}) {
    const aad = aadFromContext(aadContext);
    const payload = deserializeEncryptedPayload(encrypted);
    return decryptAes256Gcm({
      ...payload,
      key: this.getKey(encrypted.keyVersion),
      aad
    });
  }

  toJSON() {
    const keys = {};
    for (const [version, key] of this.keysByVersion.entries()) {
      keys[String(version)] = toBase64(key);
    }
    return {
      currentVersion: this.currentVersion,
      keys
    };
  }

  static fromJSON(value) {
    if (!value || typeof value !== "object") {
      throw new TypeError("KeyRing JSON must be an object");
    }
    const { currentVersion, keys } = value;
    if (!keys || typeof keys !== "object") {
      throw new TypeError("KeyRing JSON missing keys");
    }
    const keysByVersion = new Map();
    for (const [versionStr, keyBase64] of Object.entries(keys)) {
      const version = Number.parseInt(versionStr, 10);
      if (!Number.isInteger(version) || version < 1) {
        throw new Error(`Invalid key version: ${versionStr}`);
      }
      const decoded = fromBase64(keyBase64, `keys[${versionStr}]`);
      assertBufferLength(decoded, AES_256_KEY_BYTES, `keys[${versionStr}]`);
      keysByVersion.set(version, decoded);
    }
    return new KeyRing({ currentVersion, keysByVersion });
  }
}
