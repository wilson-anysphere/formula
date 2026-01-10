import { DatabaseSync } from "node:sqlite";
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
    this._db = null;
  }

  async _open() {
    if (this._db) return this._db;
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
    this._db = new DatabaseSync(this.filePath);
    this._db.exec(`
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
    return this._db;
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
    stmt.run(
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
    );
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const db = await this._open();
    const row = db
      .prepare(
        `SELECT id, kind, timestamp_ms, user_id, user_name, description,
                checkpoint_name, checkpoint_locked, checkpoint_annotations, snapshot
         FROM versions WHERE id = ?`
      )
      .get(versionId);
    if (!row) return null;
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
    const rows = db
      .prepare(
        `SELECT id, kind, timestamp_ms, user_id, user_name, description,
                checkpoint_name, checkpoint_locked, checkpoint_annotations, snapshot
         FROM versions ORDER BY timestamp_ms DESC`
      )
      .all();
    return rows.map((row) => ({
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
    }));
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    const db = await this._open();
    if (patch.checkpointLocked === undefined) return;
    db.prepare(`UPDATE versions SET checkpoint_locked = ? WHERE id = ?`).run(
      patch.checkpointLocked ? 1 : 0,
      versionId
    );
  }

  close() {
    if (!this._db) return;
    this._db.close();
    this._db = null;
  }
}

