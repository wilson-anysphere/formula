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
 * - When caching Arrow IPC results, the store writes a separate `.bin` blob to
 *   avoid JSON inflation (and to preserve raw bytes). The `.json` entry file
 *   contains metadata plus a marker pointing at the companion `.bin`.
 * - `enableEncryption()` / `disableEncryption()` migrate existing entries and
 *   can be expensive for large cache directories.
 */
export class EncryptedFileSystemCacheStore {
  /**
   * @param {{
   *   directory: string;
   *   now?: () => number;
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
    this.now = options.now ?? (() => Date.now());

    this._binaryMarkerKey = "__pq_cache_binary";

    /** @type {{ fs: typeof import("node:fs/promises"), path: typeof import("node:path") } | null} */
    this._deps = null;

    /** @type {KeyRing | null} */
    this._keyRing = null;
    /** @type {Promise<KeyRing | null> | null} */
    this._keyRingLoadPromise = null;
    /** @type {Promise<KeyRing> | null} */
    this._keyRingEnsurePromise = null;
  }

  /**
   * Best-effort touch to update file mtimes as an approximation of last-access time.
   * Used for LRU eviction without needing to rewrite/decrypt cache payloads.
   *
   * @param {string} filePath
   * @param {number} nowMs
   */
  async _touchFile(filePath, nowMs) {
    const { fs } = await this.deps();
    try {
      const date = new Date(nowMs);
      await fs.utimes(filePath, date, date);
    } catch {
      // ignore
    }
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
   * @returns {Promise<{ jsonPath: string, binPath: string, binFileName: string }>}
   */
  async pathsForKey(key) {
    const hashed = fnv1a64(key);
    const { path } = await this.deps();
    return {
      jsonPath: path.join(this.directory, `${hashed}.json`),
      binPath: path.join(this.directory, `${hashed}.bin`),
      binFileName: `${hashed}.bin`,
    };
  }

  /**
   * Backwards-compatible helper for callers that relied on the original JSON-only
   * implementation.
   *
   * @param {string} key
   * @returns {Promise<string>}
   */
  async filePathForKey(key) {
    const { jsonPath } = await this.pathsForKey(key);
    return jsonPath;
  }

  async ensureDir() {
    const { fs } = await this.deps();
    await fs.mkdir(this.directory, { recursive: true });
  }

  /**
   * @param {string} finalPath
   * @param {Buffer} bytes
   */
  async _atomicWrite(finalPath, bytes) {
    const { fs, path } = await this.deps();
    const tmp = path.join(
      path.dirname(finalPath),
      `${path.basename(finalPath)}.tmp-${Date.now()}-${crypto.randomBytes(8).toString("hex")}`,
    );
    try {
      await fs.writeFile(tmp, bytes);
      try {
        await fs.rename(tmp, finalPath);
      } catch (err) {
        // On Windows, rename does not reliably overwrite existing files.
        if (err && typeof err === "object" && "code" in err && (err.code === "EEXIST" || err.code === "EPERM")) {
          await fs.rm(finalPath, { force: true });
          await fs.rename(tmp, finalPath);
          return;
        }
        throw err;
      }
    } catch (err) {
      await fs.rm(tmp, { force: true }).catch(() => {});
      throw err;
    }
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

    const plaintext = await this._decryptFileBytes(bytes);
    if (!plaintext) return null;

    try {
      const parsed = JSON.parse(plaintext.toString("utf8"));
      if (!parsed || typeof parsed.key !== "string" || !parsed.entry) return null;
      return { key: parsed.key, entry: parsed.entry };
    } catch {
      return null;
    }
  }

  /**
   * @param {Buffer} bytes
   * @returns {Promise<Buffer | null>}
   */
  async _decryptFileBytes(bytes) {
    if (!isEncryptedFileBytes(bytes)) return bytes;
    if (!this.encryption) return null;
    const ring = await this._loadKeyRing();
    if (!ring) return null;
    try {
      const decoded = decodeEncryptedFileBytes(bytes);
      return ring.decryptBytes(decoded, { aadContext: this._aadContext() });
    } catch {
      return null;
    }
  }

  /**
   * @param {Buffer} plaintext
   * @returns {Promise<Buffer>}
   */
  async _encryptFileBytes(plaintext) {
    if (!this._encryptionEnabled) return plaintext;
    if (!this.encryption) throw new Error("encryption is not configured");
    const ring = await this._ensureKeyRing();
    const encrypted = ring.encryptBytes(plaintext, { aadContext: this._aadContext() });
    return encodeEncryptedFileBytes({
      keyVersion: encrypted.keyVersion,
      iv: encrypted.iv,
      tag: encrypted.tag,
      ciphertext: encrypted.ciphertext,
    });
  }

  /**
   * If the cache entry refers to a companion `.bin` blob, load it and hydrate the
   * entry payload.
   *
   * @param {string} key
   * @param {CacheEntry} entry
   * @returns {Promise<CacheEntry | null>}
   */
  async _hydrateBinary(key, entry) {
    const value = entry?.value;
    const bytesMarker = value?.version === 2 && value?.table?.kind === "arrow" ? value?.table?.bytes : null;
    if (
      !bytesMarker ||
      typeof bytesMarker !== "object" ||
      Array.isArray(bytesMarker) ||
      typeof bytesMarker[this._binaryMarkerKey] !== "string"
    ) {
      return entry;
    }

    const { fs } = await this.deps();
    const { binPath, binFileName } = await this.pathsForKey(key);
    const markerFileName = bytesMarker[this._binaryMarkerKey];
    // Only allow loading from within the cache directory.
    if (markerFileName !== binFileName) return null;

    /** @type {Buffer} */
    let bytes;
    try {
      bytes = await fs.readFile(binPath);
    } catch {
      return null;
    }

    const plaintext = await this._decryptFileBytes(bytes);
    if (!plaintext) return null;

    const restored = new Uint8Array(plaintext.buffer, plaintext.byteOffset, plaintext.byteLength);
    return {
      ...entry,
      value: {
        ...value,
        table: {
          ...value.table,
          bytes: restored,
        },
      },
    };
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    await this.ensureDir();
    const { jsonPath, binPath } = await this.pathsForKey(key);
    const parsed = await this._readAndParseFile(jsonPath);
    if (!parsed || parsed.key !== key) return null;

    const value = parsed.entry?.value;
    const bytesMarker = value?.version === 2 && value?.table?.kind === "arrow" ? value?.table?.bytes : null;
    const hasBin =
      bytesMarker &&
      typeof bytesMarker === "object" &&
      !Array.isArray(bytesMarker) &&
      // @ts-ignore - runtime access
      typeof bytesMarker[this._binaryMarkerKey] === "string";

    const hydrated = await this._hydrateBinary(key, parsed.entry);
    if (!hydrated) return null;

    const nowMs = this.now();
    await this._touchFile(jsonPath, nowMs);
    if (hasBin) {
      await this._touchFile(binPath, nowMs);
    }

    return hydrated;
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    await this.ensureDir();
    const { fs } = await this.deps();
    const { jsonPath, binPath, binFileName } = await this.pathsForKey(key);

    const value = entry?.value;
    const arrowBytes =
      value?.version === 2 && value?.table?.kind === "arrow" && value?.table?.bytes instanceof Uint8Array
        ? value.table.bytes
        : null;

    if (arrowBytes) {
      const binPlaintext = Buffer.from(arrowBytes.buffer, arrowBytes.byteOffset, arrowBytes.byteLength);
      const binOut = await this._encryptFileBytes(binPlaintext);
      await this._atomicWrite(binPath, binOut);

      const patchedValue = {
        ...value,
        table: {
          ...value.table,
          bytes: { [this._binaryMarkerKey]: binFileName },
        },
      };

      const jsonPlaintext = Buffer.from(JSON.stringify({ key, entry: { ...entry, value: patchedValue } }), "utf8");
      const jsonOut = await this._encryptFileBytes(jsonPlaintext);
      await this._atomicWrite(jsonPath, jsonOut);

      const nowMs = this.now();
      await this._touchFile(binPath, nowMs);
      await this._touchFile(jsonPath, nowMs);
      return;
    }

    // If we are writing a JSON-only entry, clean up any previous binary blob.
    await fs.rm(binPath, { force: true });
    const jsonPlaintext = Buffer.from(JSON.stringify({ key, entry }), "utf8");
    const jsonOut = await this._encryptFileBytes(jsonPlaintext);
    await this._atomicWrite(jsonPath, jsonOut);

    await this._touchFile(jsonPath, this.now());
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    const { fs } = await this.deps();
    const { jsonPath, binPath } = await this.pathsForKey(key);
    await fs.rm(jsonPath, { force: true });
    await fs.rm(binPath, { force: true });
  }

  async clear() {
    const { fs } = await this.deps();
    await fs.rm(this.directory, { recursive: true, force: true });
  }

  /**
   * Proactively delete expired entries.
   *
   * CacheManager deletes expired keys on access, but long-lived caches can benefit
   * from an occasional sweep to free disk space (including `.bin` blobs).
   *
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs = Date.now()) {
    await this.ensureDir();
    const { fs, path } = await this.deps();
    const tmpGraceMs = 5 * 60 * 1000;
    const orphanGraceMs = tmpGraceMs;

    /** @type {import("node:fs").Dirent[]} */
    let entries = [];
    try {
      entries = await fs.readdir(this.directory, { withFileTypes: true });
    } catch {
      return;
    }

    /** @type {Set<string>} */
    const liveJsonBases = new Set();

    for (const entry of entries) {
      if (!entry.isFile()) continue;

      // Best-effort cleanup of stale temp files left behind by interrupted writes.
      if (entry.name.includes(".tmp-") || entry.name.endsWith(".tmp")) {
        const tmpPath = path.join(this.directory, entry.name);
        try {
          const stats = await fs.stat(tmpPath);
          const ageMs = nowMs - stats.mtimeMs;
          if (Number.isFinite(stats.mtimeMs) && ageMs > tmpGraceMs) {
            await fs.rm(tmpPath, { force: true });
          }
        } catch {
          // ignore
        }
        continue;
      }

      if (!entry.name.endsWith(".json")) continue;
      const baseName = entry.name.slice(0, -".json".length);
      const filePath = path.join(this.directory, entry.name);
      const binPath = path.join(this.directory, `${baseName}.bin`);

      /** @type {Buffer} */
      let bytes;
      try {
        bytes = await fs.readFile(filePath);
      } catch {
        continue;
      }

      const plaintext = await this._decryptFileBytes(bytes);
      if (!plaintext) {
        // Best-effort cleanup: if we can't decrypt/parse the entry, treat it as
        // corrupted cache data and remove it.
        await fs.rm(filePath, { force: true }).catch(() => {});
        await fs.rm(binPath, { force: true }).catch(() => {});
        continue;
      }

      try {
        const parsed = JSON.parse(plaintext.toString("utf8"));
        const expiresAtMs =
          parsed && typeof parsed === "object" && parsed.entry && typeof parsed.entry === "object" ? parsed.entry.expiresAtMs : null;
        if (typeof expiresAtMs === "number" && expiresAtMs <= nowMs) {
          await fs.rm(filePath, { force: true });
          await fs.rm(binPath, { force: true });
        } else {
          liveJsonBases.add(baseName);
        }
      } catch {
        // Corrupted or unreadable cache entries are treated as misses; remove them
        // so they don't linger indefinitely.
        await fs.rm(filePath, { force: true }).catch(() => {});
        await fs.rm(binPath, { force: true }).catch(() => {});
      }
    }

    // Best-effort cleanup of orphaned `.bin` blobs left behind by interrupted writes.
    for (const entry of entries) {
      if (!entry.isFile()) continue;
      if (!entry.name.endsWith(".bin")) continue;
      const base = entry.name.slice(0, -".bin".length);
      if (!/^[0-9a-f]{16}$/.test(base)) continue;
      if (liveJsonBases.has(base)) continue;

      const binPath = path.join(this.directory, entry.name);
      try {
        const stats = await fs.stat(binPath);
        const ageMs = nowMs - stats.mtimeMs;
        if (Number.isFinite(stats.mtimeMs) && ageMs > orphanGraceMs) {
          await fs.rm(binPath, { force: true });
        }
      } catch {
        // ignore
      }
    }
  }

  /**
   * Prune expired entries and enforce optional entry/byte quotas using LRU eviction.
   *
   * This store approximates LRU using the entry file's mtime (updated on `get()` and
   * `set()`).
   *
   * @param {{ nowMs: number, maxEntries?: number, maxBytes?: number }} options
   */
  async prune(options) {
    const maxEntries = options.maxEntries;
    const maxBytes = options.maxBytes;

    if (maxEntries == null && maxBytes == null) {
      await this.pruneExpired(options.nowMs);
      return;
    }

    await this.ensureDir();
    const { fs, path } = await this.deps();

    /** @type {import("node:fs").Dirent[]} */
    let entries = [];
    try {
      entries = await fs.readdir(this.directory, { withFileTypes: true });
    } catch {
      return;
    }

    /** @type {Array<{ base: string, jsonPath: string, binPath: string, lastAccessMs: number, sizeBytes: number }>} */
    const records = [];

    for (const entry of entries) {
      if (!entry.isFile()) continue;

      // Ignore temp files (pruneExpired handles stale cleanup).
      if (entry.name.includes(".tmp-") || entry.name.endsWith(".tmp")) continue;

      if (!entry.name.endsWith(".json")) continue;
      const baseName = entry.name.slice(0, -".json".length);
      const jsonPath = path.join(this.directory, entry.name);
      const binPath = path.join(this.directory, `${baseName}.bin`);

      let jsonStat;
      try {
        jsonStat = await fs.stat(jsonPath);
      } catch {
        continue;
      }

      let binSize = 0;
      try {
        const binStat = await fs.stat(binPath);
        binSize = binStat.size;
      } catch {
        binSize = 0;
      }

      records.push({
        base: baseName,
        jsonPath,
        binPath,
        lastAccessMs: Number.isFinite(jsonStat.mtimeMs) ? jsonStat.mtimeMs : 0,
        sizeBytes: jsonStat.size + binSize,
      });
    }

    let totalEntries = records.length;
    let totalBytes = 0;
    for (const rec of records) totalBytes += rec.sizeBytes;

    records.sort((a, b) => {
      if (a.lastAccessMs !== b.lastAccessMs) return a.lastAccessMs - b.lastAccessMs;
      return a.base.localeCompare(b.base);
    });

    let idx = 0;
    while (
      (maxEntries != null && totalEntries > maxEntries) ||
      (maxBytes != null && totalBytes > maxBytes)
    ) {
      const victim = records[idx++];
      if (!victim) break;
      await fs.rm(victim.jsonPath, { force: true }).catch(() => {});
      await fs.rm(victim.binPath, { force: true }).catch(() => {});
      totalEntries -= 1;
      totalBytes -= victim.sizeBytes;
    }
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
      return entries
        .filter((e) => e.isFile() && e.name.endsWith(".json"))
        .map((e) => path.join(this.directory, e.name));
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
      const hydrated = await this._hydrateBinary(parsed.key, parsed.entry);
      if (hydrated) {
        await this.set(parsed.key, hydrated);
      }
      if (expectedPath !== filePath) {
        await fs.rm(filePath, { force: true });
        await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
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
        await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
        continue;
      }

      try {
        const decoded = decodeEncryptedFileBytes(bytes);
        const plaintext = ring.decryptBytes(decoded, { aadContext: this._aadContext() });
        const parsed = JSON.parse(plaintext.toString("utf8"));
        if (!parsed || typeof parsed.key !== "string" || !parsed.entry) {
          await fs.rm(filePath, { force: true });
          // Best-effort cleanup of any companion blob.
          await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
          continue;
        }

        const expectedPath = await this.filePathForKey(parsed.key);
        const hydrated = await this._hydrateBinary(parsed.key, parsed.entry);
        if (hydrated) {
          await this.set(parsed.key, hydrated);
        } else {
          await fs.rm(filePath, { force: true });
          await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
          continue;
        }
        if (expectedPath !== filePath) {
          await fs.rm(filePath, { force: true });
          await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
        }
      } catch {
        await fs.rm(filePath, { force: true });
        await fs.rm(filePath.replace(/\.json$/, ".bin"), { force: true }).catch(() => {});
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
