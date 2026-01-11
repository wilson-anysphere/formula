import crypto from "node:crypto";

import { fnv1a64 } from "./key.js";

import { KeyRing } from "../../../security/crypto/keyring.js";
import {
  decodeEncryptedFileBytes,
  encodeEncryptedFileBytes,
  isEncryptedFileBytes
} from "../../../security/crypto/encryptedFile.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

/**
 * Encrypted filesystem cache store for Node environments.
 *
 * This is a drop-in replacement for `FileSystemCacheStore` that optionally
 * encrypts cache entries at rest using AES-256-GCM via `KeyRing`.
 *
 * Notes:
 * - This store is tolerant of mixed-mode directories (some plaintext JSON files,
 *   some encrypted blobs) to support migration.
 * - `enableEncryption()` / `disableEncryption()` migrate existing entries and
 *   can be expensive for large cache directories.
 */
export class EncryptedFileSystemCacheStore {
  /**
   * @param {{
   *   directory: string;
   *   encryption: {
   *     enabled: boolean;
   *     keychainProvider: {
   *       getSecret: (opts: { service: string; account: string }) => Promise<Buffer | null>;
   *       setSecret: (opts: { service: string; account: string; secret: Buffer }) => Promise<void>;
   *       deleteSecret: (opts: { service: string; account: string }) => Promise<void>;
   *     };
   *     keychainService?: string;
   *     keychainAccount?: string;
   *     aadContext?: any;
   *   };
   * }} options
   */
  constructor(options) {
    this.directory = options.directory;
    this.encryption = options.encryption ?? null;
    this._encryptionEnabled = Boolean(options.encryption?.enabled);

    /** @type {{ fs: typeof import("node:fs/promises"), path: typeof import("node:path") } | null} */
    this._deps = null;

    /** @type {KeyRing | null} */
    this._keyRing = null;
    /** @type {Promise<KeyRing | null> | null} */
    this._keyRingLoadPromise = null;
    /** @type {Promise<KeyRing> | null} */
    this._keyRingEnsurePromise = null;
  }

  _aadContext() {
    return this.encryption?.aadContext ?? { scope: "formula.power-query.cache.fs", schemaVersion: 1 };
  }

  _keychainService() {
    return this.encryption?.keychainService ?? "formula.desktop";
  }

  _keychainAccount() {
    if (this.encryption?.keychainAccount) return this.encryption.keychainAccount;
    const hash = crypto.createHash("sha256").update(this.directory).digest("hex").slice(0, 16);
    return `power-query-cache:${hash}`;
  }

  /**
   * @returns {Promise<KeyRing | null>}
   */
  async _loadKeyRing() {
    if (this._keyRing) return this._keyRing;
    if (!this.encryption) return null;
    if (this._keyRingLoadPromise) return this._keyRingLoadPromise;

    this._keyRingLoadPromise = (async () => {
      const secret = await this.encryption.keychainProvider.getSecret({
        service: this._keychainService(),
        account: this._keychainAccount()
      });
      if (!secret) return null;
      const parsed = JSON.parse(secret.toString("utf8"));
      const ring = KeyRing.fromJSON(parsed);
      this._keyRing = ring;
      return ring;
    })();

    try {
      return await this._keyRingLoadPromise;
    } finally {
      this._keyRingLoadPromise = null;
    }
  }

  /**
   * @param {KeyRing} keyRing
   */
  async _storeKeyRing(keyRing) {
    if (!this.encryption) throw new Error("encryption is not configured");
    const json = JSON.stringify(keyRing.toJSON());
    await this.encryption.keychainProvider.setSecret({
      service: this._keychainService(),
      account: this._keychainAccount(),
      secret: Buffer.from(json, "utf8")
    });
    this._keyRing = keyRing;
  }

  async _deleteKeyRing() {
    if (!this.encryption) return;
    await this.encryption.keychainProvider.deleteSecret({
      service: this._keychainService(),
      account: this._keychainAccount()
    });
    this._keyRing = null;
  }

  /**
   * Ensure a key ring exists in the keychain (creating one if missing).
   *
   * This method is concurrency-safe: multiple concurrent calls will share a
   * single key generation + keychain write.
   *
   * @returns {Promise<KeyRing>}
   */
  async _ensureKeyRing() {
    if (!this.encryption) throw new Error("encryption is not configured");
    const existing = await this._loadKeyRing();
    if (existing) return existing;
    if (this._keyRingEnsurePromise) return this._keyRingEnsurePromise;

    this._keyRingEnsurePromise = (async () => {
      const created = KeyRing.create();
      await this._storeKeyRing(created);
      return created;
    })();

    try {
      return await this._keyRingEnsurePromise;
    } finally {
      this._keyRingEnsurePromise = null;
    }
  }

  /**
   * @returns {Promise<{ fs: typeof import("node:fs/promises"), path: typeof import("node:path") }>}
   */
  async deps() {
    if (this._deps) return this._deps;
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    this._deps = { fs, path };
    return this._deps;
  }

  /**
   * @param {string} key
   * @returns {Promise<string>}
   */
  async filePathForKey(key) {
    const hashed = fnv1a64(key);
    const { path } = await this.deps();
    return path.join(this.directory, `${hashed}.json`);
  }

  async ensureDir() {
    const { fs } = await this.deps();
    await fs.mkdir(this.directory, { recursive: true });
  }

  /**
   * @param {string} filePath
   * @param {Buffer} bytes
   */
  async _atomicWrite(filePath, bytes) {
    const { fs } = await this.deps();
    const tmp = `${filePath}.${crypto.randomBytes(8).toString("hex")}.tmp`;
    await fs.writeFile(tmp, bytes);
    await fs.rename(tmp, filePath);
  }

  /**
   * @param {string} filePath
   * @returns {Promise<{ key: string, entry: CacheEntry } | null>}
   */
  async _readAndParseFile(filePath) {
    const { fs } = await this.deps();
    /** @type {Buffer} */
    let bytes;
    try {
      bytes = await fs.readFile(filePath);
    } catch {
      return null;
    }

    try {
      if (isEncryptedFileBytes(bytes)) {
        if (!this.encryption) return null;
        const ring = await this._loadKeyRing();
        if (!ring) return null;
        const decoded = decodeEncryptedFileBytes(bytes);
        const plaintext = ring.decryptBytes(decoded, { aadContext: this._aadContext() });
        const parsed = JSON.parse(plaintext.toString("utf8"));
        if (!parsed || typeof parsed.key !== "string" || !parsed.entry) return null;
        return { key: parsed.key, entry: parsed.entry };
      }

      const parsed = JSON.parse(bytes.toString("utf8"));
      if (!parsed || typeof parsed.key !== "string" || !parsed.entry) return null;
      return { key: parsed.key, entry: parsed.entry };
    } catch {
      return null;
    }
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    await this.ensureDir();
    const filePath = await this.filePathForKey(key);
    const parsed = await this._readAndParseFile(filePath);
    if (!parsed || parsed.key !== key) return null;
    return parsed.entry ?? null;
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    await this.ensureDir();
    const filePath = await this.filePathForKey(key);
    const payload = Buffer.from(JSON.stringify({ key, entry }), "utf8");
    let out = payload;

    if (this._encryptionEnabled) {
      if (!this.encryption) throw new Error("encryption is not configured");
      const ring = await this._ensureKeyRing();
      const encrypted = ring.encryptBytes(payload, { aadContext: this._aadContext() });
      out = encodeEncryptedFileBytes({
        keyVersion: encrypted.keyVersion,
        iv: encrypted.iv,
        tag: encrypted.tag,
        ciphertext: encrypted.ciphertext
      });
    }

    await this._atomicWrite(filePath, out);
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    const { fs } = await this.deps();
    const filePath = await this.filePathForKey(key);
    await fs.rm(filePath, { force: true });
  }

  async clear() {
    const { fs } = await this.deps();
    await fs.rm(this.directory, { recursive: true, force: true });
  }

  /**
   * List entry files in the cache directory.
   *
   * @returns {Promise<string[]>}
   */
  async _listEntryFiles() {
    const { fs, path } = await this.deps();
    try {
      const entries = await fs.readdir(this.directory, { withFileTypes: true });
      return entries.filter((e) => e.isFile()).map((e) => path.join(this.directory, e.name));
    } catch {
      return [];
    }
  }

  /**
   * Enable encryption for new writes and migrate existing plaintext entries in
   * the directory to encrypted format.
   */
  async enableEncryption() {
    if (!this.encryption) throw new Error("encryption is not configured");
    this._encryptionEnabled = true;
    await this._ensureKeyRing();

    const { fs } = await this.deps();
    const files = await this._listEntryFiles();
    for (const filePath of files) {
      /** @type {Buffer} */
      let bytes;
      try {
        bytes = await fs.readFile(filePath);
      } catch {
        continue;
      }
      if (isEncryptedFileBytes(bytes)) continue;

      let parsed;
      try {
        parsed = JSON.parse(bytes.toString("utf8"));
      } catch {
        continue;
      }
      if (!parsed || typeof parsed.key !== "string" || !parsed.entry) continue;

      const expectedPath = await this.filePathForKey(parsed.key);
      await this.set(parsed.key, parsed.entry);
      if (expectedPath !== filePath) {
        await fs.rm(filePath, { force: true });
      }
    }
  }

  /**
   * Disable encryption for new writes and migrate encrypted entries in the
   * directory to plaintext JSON.
   *
   * @param {{ deleteKey?: boolean }} [options]
   */
  async disableEncryption({ deleteKey = true } = {}) {
    this._encryptionEnabled = false;

    const { fs } = await this.deps();
    const files = await this._listEntryFiles();
    const ring = await this._loadKeyRing();

    for (const filePath of files) {
      /** @type {Buffer} */
      let bytes;
      try {
        bytes = await fs.readFile(filePath);
      } catch {
        continue;
      }
      if (!isEncryptedFileBytes(bytes)) continue;

      if (!this.encryption || !ring) {
        // Cache data is optional; if we can't decrypt, drop the entry so we don't
        // leave behind unreadable encrypted blobs.
        await fs.rm(filePath, { force: true });
        continue;
      }

      try {
        const decoded = decodeEncryptedFileBytes(bytes);
        const plaintext = ring.decryptBytes(decoded, { aadContext: this._aadContext() });
        const parsed = JSON.parse(plaintext.toString("utf8"));
        if (!parsed || typeof parsed.key !== "string" || !parsed.entry) {
          await fs.rm(filePath, { force: true });
          continue;
        }

        const expectedPath = await this.filePathForKey(parsed.key);
        await this.set(parsed.key, parsed.entry);
        if (expectedPath !== filePath) {
          await fs.rm(filePath, { force: true });
        }
      } catch {
        await fs.rm(filePath, { force: true });
      }
    }

    if (deleteKey) {
      await this._deleteKeyRing();
    }
  }

  /**
   * Rotate the active encryption key used for new writes.
   *
   * Existing entries remain decryptable because the key ring retains older
   * versions; to re-encrypt all existing entries you must rewrite them (e.g. by
   * clearing the cache or implementing a full migration step).
   */
  async rotateKey() {
    if (!this.encryption) throw new Error("encryption is not configured");
    if (!this._encryptionEnabled) {
      throw new Error("Cannot rotate key: encryption is disabled");
    }
    const ring = await this._ensureKeyRing();
    ring.rotate();
    await this._storeKeyRing(ring);
  }
}

