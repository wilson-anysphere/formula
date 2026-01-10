import fs from "node:fs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";

/**
 * SQLite-backed audit log store.
 *
 * Note: Node's `node:sqlite` is currently marked experimental, but it avoids
 * pulling in native npm dependencies and is sufficient for a baseline audit
 * trail in local development and CI.
 */
export class SqliteAuditLogStore {
  /**
   * @param {object} options
   * @param {string} options.path - SQLite database path. Use ":memory:" for tests.
   */
  constructor({ path: dbPath }) {
    if (typeof dbPath !== "string" || dbPath.length === 0) {
      throw new TypeError("SqliteAuditLogStore requires a database path");
    }

    this.path = dbPath;

    if (dbPath !== ":memory:") {
      const dir = path.dirname(dbPath);
      fs.mkdirSync(dir, { recursive: true });
    }

    this.db = new DatabaseSync(dbPath);
    this.db.exec("PRAGMA journal_mode = WAL;");
    this.db.exec("PRAGMA foreign_keys = ON;");

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS audit_events (
        id TEXT PRIMARY KEY,
        ts INTEGER NOT NULL,
        event_type TEXT NOT NULL,
        actor_type TEXT NOT NULL,
        actor_id TEXT NOT NULL,
        success INTEGER NOT NULL,
        metadata TEXT NOT NULL
      );
    `);

    this.db.exec(`CREATE INDEX IF NOT EXISTS idx_audit_events_ts ON audit_events(ts);`);
    this.db.exec(
      `CREATE INDEX IF NOT EXISTS idx_audit_events_actor ON audit_events(actor_type, actor_id);`
    );

    this.insertStmt = this.db.prepare(
      `INSERT INTO audit_events (id, ts, event_type, actor_type, actor_id, success, metadata)
       VALUES (?, ?, ?, ?, ?, ?, ?);`
    );

    this.queryStmt = this.db.prepare(
      `SELECT id, ts, event_type, actor_type, actor_id, success, metadata
         FROM audit_events
        WHERE (?1 IS NULL OR actor_type = ?1)
          AND (?2 IS NULL OR actor_id = ?2)
          AND (?3 IS NULL OR event_type = ?3)
          AND (?4 IS NULL OR ts >= ?4)
          AND (?5 IS NULL OR ts <= ?5)
        ORDER BY ts DESC
        LIMIT COALESCE(?6, 500);`
    );
  }

  /**
   * @param {object} event
   * @param {string} event.id
   * @param {number} event.ts
   * @param {string} event.eventType
   * @param {{type: string, id: string}} event.actor
   * @param {boolean} event.success
   * @param {object} event.metadata
   */
  append(event) {
    this.insertStmt.run(
      event.id,
      event.ts,
      event.eventType,
      event.actor.type,
      event.actor.id,
      event.success ? 1 : 0,
      JSON.stringify(event.metadata ?? {})
    );
  }

  /**
   * @param {object} filters
   * @param {string | null} [filters.actorType]
   * @param {string | null} [filters.actorId]
   * @param {string | null} [filters.eventType]
   * @param {number | null} [filters.startTs]
   * @param {number | null} [filters.endTs]
   * @param {number | null} [filters.limit]
   */
  query(filters = {}) {
    const rows = this.queryStmt.all(
      filters.actorType ?? null,
      filters.actorId ?? null,
      filters.eventType ?? null,
      filters.startTs ?? null,
      filters.endTs ?? null,
      filters.limit ?? null
    );

    return rows.map((row) => ({
      id: row.id,
      ts: row.ts,
      eventType: row.event_type,
      actor: { type: row.actor_type, id: row.actor_id },
      success: Boolean(row.success),
      metadata: JSON.parse(row.metadata)
    }));
  }
}
