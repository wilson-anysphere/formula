import initSqlJs from "sql.js";
import { createRequire } from "node:module";
import { promises as fs } from "node:fs";
import path from "node:path";

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
 * keeping the implementation dependency-free by using Node's built-in
 * `node:sqlite` module.
 */
export class SQLiteVersionStore {
  /**
   * @param {{ filePath: string }} opts
   */
  constructor(opts) {
    this.filePath = opts.filePath;
    /** @type {any | null} */
    this._db = null;
    /** @type {Promise<any> | null} */
    this._initPromise = null;
    /** @type {Promise<void>} */
    this._persistChain = Promise.resolve();
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

    const db = existing ? new SQL.Database(existing) : new SQL.Database();
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
    const data = db.export();
    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, data);
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

  close() {
    if (!this._db) return;
    this._db.close();
    this._db = null;
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
