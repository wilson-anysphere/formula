import fs from "node:fs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";

import {
  AUDIT_EVENT_SCHEMA_VERSION,
  assertAuditEvent,
  auditEventToSqliteRow,
  createAuditEvent,
  retentionCutoffMs
} from "../../../audit-core/index.js";

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

    this.#ensureSchema();

    this.insertStmt = this.db.prepare(
      `INSERT INTO audit_events (
         id,
         ts,
         timestamp,
         event_type,
         actor_type,
         actor_id,
         org_id,
         user_id,
         user_email,
         ip_address,
         user_agent,
         session_id,
         resource_type,
         resource_id,
         resource_name,
         success,
         error_code,
         error_message,
         details,
         request_id,
         trace_id
       )
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);`
    );

    this.queryStmt = this.db.prepare(
      `SELECT
          id,
          ts,
          timestamp,
          event_type,
          actor_type,
          actor_id,
          org_id,
          user_id,
          user_email,
          ip_address,
          user_agent,
          session_id,
          resource_type,
          resource_id,
          resource_name,
          success,
          error_code,
          error_message,
          details,
          request_id,
          trace_id
         FROM audit_events
        WHERE (?1 IS NULL OR actor_type = ?1)
          AND (?2 IS NULL OR actor_id = ?2)
          AND (?3 IS NULL OR event_type = ?3)
          AND (?4 IS NULL OR ts >= ?4)
          AND (?5 IS NULL OR ts <= ?5)
        ORDER BY ts DESC
        LIMIT COALESCE(?6, 500);`
    );

    this.retentionStmt = this.db.prepare(`DELETE FROM audit_events WHERE ts < ?;`);
  }

  #ensureSchema() {
    const userVersionRow = this.db.prepare("PRAGMA user_version").get();
    const userVersion = Number(userVersionRow?.user_version ?? 0);

    // Fresh DB.
    const columns = this.db.prepare("PRAGMA table_info(audit_events)").all();
    if (columns.length === 0) {
      this.#createSchemaV1();
      this.db.exec(`PRAGMA user_version = ${AUDIT_EVENT_SCHEMA_VERSION};`);
      return;
    }

    const hasLegacyMetadata = columns.some((col) => col.name === "metadata");
    const hasCanonicalTimestamp = columns.some((col) => col.name === "timestamp");
    const shouldMigrate = hasLegacyMetadata || !hasCanonicalTimestamp;

    if (!shouldMigrate) {
      if (userVersion < AUDIT_EVENT_SCHEMA_VERSION) {
        this.db.exec(`PRAGMA user_version = ${AUDIT_EVENT_SCHEMA_VERSION};`);
      }
      return;
    }

    this.db.exec("BEGIN;");
    try {
      this.#createSchemaV1({ tableName: "audit_events_new" });

      const legacyRows = this.db
        .prepare(`SELECT id, ts, event_type, actor_type, actor_id, success, metadata FROM audit_events`)
        .all();

      const insertNew = this.db.prepare(
        `INSERT INTO audit_events_new (
           id,
           ts,
           timestamp,
           event_type,
           actor_type,
           actor_id,
           org_id,
           user_id,
           user_email,
           ip_address,
           user_agent,
           session_id,
           resource_type,
           resource_id,
           resource_name,
           success,
           error_code,
           error_message,
           details,
           request_id,
           trace_id
         )
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);`
      );

      for (const row of legacyRows) {
        const migrated = createAuditEvent({
          id: row.id,
          timestamp: new Date(row.ts).toISOString(),
          eventType: row.event_type,
          actor: { type: row.actor_type, id: row.actor_id },
          success: Boolean(row.success),
          details: row.metadata ? JSON.parse(row.metadata) : {}
        });
        const sqliteRow = auditEventToSqliteRow(migrated);
        insertNew.run(
          sqliteRow.id,
          sqliteRow.ts,
          sqliteRow.timestamp,
          sqliteRow.eventType,
          sqliteRow.actorType,
          sqliteRow.actorId,
          sqliteRow.orgId,
          sqliteRow.userId,
          sqliteRow.userEmail,
          sqliteRow.ipAddress,
          sqliteRow.userAgent,
          sqliteRow.sessionId,
          sqliteRow.resourceType,
          sqliteRow.resourceId,
          sqliteRow.resourceName,
          sqliteRow.success,
          sqliteRow.errorCode,
          sqliteRow.errorMessage,
          sqliteRow.details,
          sqliteRow.requestId,
          sqliteRow.traceId
        );
      }

      this.db.exec("DROP TABLE audit_events;");
      this.db.exec("ALTER TABLE audit_events_new RENAME TO audit_events;");
      this.#createIndexes();
      this.db.exec(`PRAGMA user_version = ${AUDIT_EVENT_SCHEMA_VERSION};`);
      this.db.exec("COMMIT;");
    } catch (error) {
      this.db.exec("ROLLBACK;");
      throw error;
    }
  }

  #createSchemaV1({ tableName = "audit_events" } = {}) {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS ${tableName} (
        id TEXT PRIMARY KEY,
        ts INTEGER NOT NULL,
        timestamp TEXT NOT NULL,
        event_type TEXT NOT NULL,
        actor_type TEXT NOT NULL,
        actor_id TEXT NOT NULL,
        org_id TEXT,
        user_id TEXT,
        user_email TEXT,
        ip_address TEXT,
        user_agent TEXT,
        session_id TEXT,
        resource_type TEXT,
        resource_id TEXT,
        resource_name TEXT,
        success INTEGER NOT NULL,
        error_code TEXT,
        error_message TEXT,
        details TEXT NOT NULL,
        request_id TEXT,
        trace_id TEXT
      );
    `);

    if (tableName === "audit_events") this.#createIndexes();
  }

  #createIndexes() {
    this.db.exec(`CREATE INDEX IF NOT EXISTS idx_audit_events_ts ON audit_events(ts);`);
    this.db.exec(
      `CREATE INDEX IF NOT EXISTS idx_audit_events_actor ON audit_events(actor_type, actor_id);`
    );
    this.db.exec(`CREATE INDEX IF NOT EXISTS idx_audit_events_event_type ON audit_events(event_type);`);
    this.db.exec(`CREATE INDEX IF NOT EXISTS idx_audit_events_org ON audit_events(org_id);`);
  }

  /**
   * @param {import("../../../audit-core/index.js").AuditEvent} event
   */
  append(event) {
    assertAuditEvent(event);
    const row = auditEventToSqliteRow(event);
    this.insertStmt.run(
      row.id,
      row.ts,
      row.timestamp,
      row.eventType,
      row.actorType,
      row.actorId,
      row.orgId,
      row.userId,
      row.userEmail,
      row.ipAddress,
      row.userAgent,
      row.sessionId,
      row.resourceType,
      row.resourceId,
      row.resourceName,
      row.success,
      row.errorCode,
      row.errorMessage,
      row.details,
      row.requestId,
      row.traceId
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

    return rows.map((row) => {
      const context = {
        orgId: row.org_id,
        userId: row.user_id,
        userEmail: row.user_email,
        ipAddress: row.ip_address,
        userAgent: row.user_agent,
        sessionId: row.session_id
      };
      const resource = row.resource_type
        ? { type: row.resource_type, id: row.resource_id, name: row.resource_name }
        : undefined;
      const error = row.error_code || row.error_message ? { code: row.error_code, message: row.error_message } : undefined;
      const correlation =
        row.request_id || row.trace_id ? { requestId: row.request_id, traceId: row.trace_id } : undefined;

      return {
        schemaVersion: AUDIT_EVENT_SCHEMA_VERSION,
        id: row.id,
        timestamp: row.timestamp,
        eventType: row.event_type,
        actor: { type: row.actor_type, id: row.actor_id },
        context,
        resource,
        success: Boolean(row.success),
        error,
        details: JSON.parse(row.details),
        correlation
      };
    });
  }

  /**
   * Delete events older than the org's configured retention window.
   *
   * @param {{ retentionDays: number, now?: number }} options
   * @returns {number} deleted row count
   */
  sweepRetention({ retentionDays, now = Date.now() }) {
    const cutoff = retentionCutoffMs(now, retentionDays);
    const result = this.retentionStmt.run(cutoff);
    return result.changes ?? 0;
  }
}
