import initSqlJs from "sql.js";
import { createRequire } from "node:module";
import crypto from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";

import { KeyRing } from "../../../security/crypto/keyring.js";
import {
  decodeEncryptedFileBytes,
  encodeEncryptedFileBytes,
  isEncryptedFileBytes
} from "../../../security/crypto/encryptedFile.js";

/**
 * @typedef {"snapshot" | "checkpoint" | "restore"} VersionKind
 *
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId: string | null;
 *   userName: string | null;
 *   description: string | null;
 *   checkpointName: string | null;
 *   checkpointLocked: boolean | null;
 *   checkpointAnnotations: string | null;
 *   snapshot: Uint8Array;
 * }} VersionRecord
 */

/**
 * SQLite-backed version store.
 *
 * This aligns with the intended desktop persistence model (SQLite) while
 * keeping the implementation dependency-light and Node>=18 compatible via
 * `sql.js` (WASM SQLite) with file persistence.
 */
export class SQLiteVersionStore {
  /**
   * @param {{
   *   filePath: string;
   *   encryption?: {
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
   * }} opts
   */
  constructor(opts) {
    this.filePath = opts.filePath;
    this.encryption = opts.encryption ?? null;
    this._encryptionEnabled = Boolean(opts.encryption?.enabled);
    /** @type {any | null} */
    this._db = null;
    /** @type {Promise<any> | null} */
    this._initPromise = null;
    /** @type {Promise<void>} */
    this._persistChain = Promise.resolve();
    /** @type {KeyRing | null} */
    this._keyRing = null;
  }

  _aadContext() {
    return this.encryption?.aadContext ?? { scope: "formula.versioning.sqlite", schemaVersion: 1 };
  }

  _keychainService() {
    return this.encryption?.keychainService ?? "formula.desktop";
  }

  _keychainAccount() {
    if (this.encryption?.keychainAccount) return this.encryption.keychainAccount;
    const hash = crypto.createHash("sha256").update(this.filePath).digest("hex").slice(0, 16);
    return `sqlite-version-store:${hash}`;
  }

  async _loadKeyRing() {
    if (this._keyRing) return this._keyRing;
    if (!this.encryption) return null;
    const secret = await this.encryption.keychainProvider.getSecret({
      service: this._keychainService(),
      account: this._keychainAccount()
    });
    if (!secret) return null;
    const parsed = JSON.parse(secret.toString("utf8"));
    this._keyRing = KeyRing.fromJSON(parsed);
    return this._keyRing;
  }

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

  async _ensureKeyRing() {
    const existing = await this._loadKeyRing();
    if (existing) return existing;
    const created = KeyRing.create();
    await this._storeKeyRing(created);
    return created;
  }

  async _open() {
    if (this._db) return this._db;
    if (this._initPromise) return this._initPromise;
    this._initPromise = this._openInner();
    return this._initPromise;
  }

  async _openInner() {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });

    const SQL = await initSqlJs({ locateFile: locateSqlJsFile });

    /** @type {Uint8Array | null} */
    let existing = null;
    try {
      existing = await fs.readFile(this.filePath);
    } catch {
      existing = null;
    }

    let bytes = existing ? Buffer.from(existing) : null;
    if (bytes && isEncryptedFileBytes(bytes)) {
      if (!this.encryption) {
        throw new Error("Encrypted SQLiteVersionStore requires encryption configuration");
      }
      const ring = await this._loadKeyRing();
      if (!ring) {
        throw new Error("Encrypted SQLiteVersionStore is missing key material in keychain");
      }
      const decoded = decodeEncryptedFileBytes(bytes);
      bytes = ring.decryptBytes(decoded, { aadContext: this._aadContext() });
    }

    const db = bytes ? new SQL.Database(bytes) : new SQL.Database();
    this._db = db;
    this._ensureSchema();
    return db;
  }

  _ensureSchema() {
    if (!this._db) return;
    this._db.run(`
      CREATE TABLE IF NOT EXISTS versions (
        id TEXT PRIMARY KEY,
        kind TEXT NOT NULL,
        timestamp_ms INTEGER NOT NULL,
        user_id TEXT,
        user_name TEXT,
        description TEXT,
        checkpoint_name TEXT,
        checkpoint_locked INTEGER,
        checkpoint_annotations TEXT,
        snapshot BLOB NOT NULL
      );

      CREATE INDEX IF NOT EXISTS idx_versions_timestamp
        ON versions(timestamp_ms DESC);
    `);
  }

  async _persist() {
    const db = await this._open();
    const data = Buffer.from(db.export());
    let out = data;

    if (this._encryptionEnabled) {
      if (!this.encryption) throw new Error("encryption is not configured");
      const ring = await this._ensureKeyRing();
      const encrypted = ring.encryptBytes(out, { aadContext: this._aadContext() });
      out = encodeEncryptedFileBytes({
        keyVersion: encrypted.keyVersion,
        iv: encrypted.iv,
        tag: encrypted.tag,
        ciphertext: encrypted.ciphertext
      });
    }

    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, this.filePath);
  }

  async _queuePersist() {
    this._persistChain = this._persistChain.then(() => this._persist());
    return this._persistChain;
  }

  /**
   * @param {VersionRecord} version
   */
  async saveVersion(version) {
    const db = await this._open();
    const stmt = db.prepare(`
      INSERT INTO versions (
        id, kind, timestamp_ms, user_id, user_name, description,
        checkpoint_name, checkpoint_locked, checkpoint_annotations, snapshot
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    stmt.run([
      version.id,
      version.kind,
      version.timestampMs,
      version.userId,
      version.userName,
      version.description,
      version.checkpointName,
      version.checkpointLocked == null ? null : version.checkpointLocked ? 1 : 0,
      version.checkpointAnnotations,
      version.snapshot
    ]);
    stmt.free();
    await this._queuePersist();
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, kind, timestamp_ms, user_id, user_name, description,
              checkpoint_name, checkpoint_locked, checkpoint_annotations, snapshot
       FROM versions WHERE id = ? LIMIT 1`
    );
    stmt.bind([versionId]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.getAsObject();
    stmt.free();
    return {
      id: row.id,
      kind: /** @type {any} */ (row.kind),
      timestampMs: row.timestamp_ms,
      userId: row.user_id ?? null,
      userName: row.user_name ?? null,
      description: row.description ?? null,
      checkpointName: row.checkpoint_name ?? null,
      checkpointLocked:
        row.checkpoint_locked == null ? null : Boolean(row.checkpoint_locked),
      checkpointAnnotations: row.checkpoint_annotations ?? null,
      snapshot: row.snapshot,
    };
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, kind, timestamp_ms, user_id, user_name, description,
              checkpoint_name, checkpoint_locked, checkpoint_annotations, snapshot
       FROM versions ORDER BY timestamp_ms DESC`
    );

    /** @type {VersionRecord[]} */
    const out = [];
    while (stmt.step()) {
      const row = stmt.getAsObject();
      out.push({
        id: row.id,
        kind: /** @type {any} */ (row.kind),
        timestampMs: row.timestamp_ms,
        userId: row.user_id ?? null,
        userName: row.user_name ?? null,
        description: row.description ?? null,
        checkpointName: row.checkpoint_name ?? null,
        checkpointLocked:
          row.checkpoint_locked == null ? null : Boolean(row.checkpoint_locked),
        checkpointAnnotations: row.checkpoint_annotations ?? null,
        snapshot: row.snapshot,
      });
    }
    stmt.free();
    return out;
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    const db = await this._open();
    if (patch.checkpointLocked === undefined) return;
    const stmt = db.prepare(`UPDATE versions SET checkpoint_locked = ? WHERE id = ?`);
    stmt.run([patch.checkpointLocked ? 1 : 0, versionId]);
    stmt.free();
    await this._queuePersist();
  }

  /**
   * @param {string} versionId
   * @returns {Promise<void>}
   */
  async deleteVersion(versionId) {
    const db = await this._open();
    const stmt = db.prepare(`DELETE FROM versions WHERE id = ?`);
    stmt.run([versionId]);
    stmt.free();
    await this._queuePersist();
  }

  close() {
    if (!this._db) return;
    this._db.close();
    this._db = null;
  }

  async enableEncryption() {
    if (!this.encryption) throw new Error("encryption is not configured");
    this._encryptionEnabled = true;
    await this._queuePersist();
  }

  async disableEncryption({ deleteKey = true } = {}) {
    this._encryptionEnabled = false;
    await this._queuePersist();
    if (deleteKey) {
      await this._deleteKeyRing();
    }
  }

  async rotateKey() {
    if (!this.encryption) throw new Error("encryption is not configured");
    if (!this._encryptionEnabled) {
      throw new Error("Cannot rotate key: encryption is disabled");
    }
    const ring = await this._ensureKeyRing();
    ring.rotate();
    await this._storeKeyRing(ring);
    await this._queuePersist();
  }
}

function locateSqlJsFile(file) {
  try {
    const require = createRequire(import.meta.url);
    return require.resolve(`sql.js/dist/${file}`);
  } catch {
    return file;
  }
}
