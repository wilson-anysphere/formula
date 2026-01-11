import type { Pool } from "pg";
import { auditLogRowToAuditEvent, type AuditEvent, type PostgresAuditLogRow } from "@formula/audit-core";

export type ExportableAuditEvent = { createdAt: Date; event: AuditEvent };

export type AuditCursor = {
  lastCreatedAt: Date | null;
  lastEventId: string | null;
};

function toDate(value: unknown): Date {
  if (value instanceof Date) return value;
  const date = new Date(String(value));
  if (Number.isNaN(date.getTime())) throw new Error(`Invalid created_at: ${String(value)}`);
  return date;
}

export async function fetchNextAuditEvents(
  db: Pool,
  orgId: string,
  cursor: AuditCursor,
  limit: number
): Promise<ExportableAuditEvent[]> {
  const hasCursor = Boolean(cursor.lastCreatedAt && cursor.lastEventId);

  const values: unknown[] = [orgId];
  if (hasCursor) {
    values.push(cursor.lastCreatedAt);
    values.push(cursor.lastEventId);
  }
  values.push(limit);

  const cursorPredicate = hasCursor
    ? "WHERE (audit_events.created_at > $2::timestamptz OR (audit_events.created_at = $2::timestamptz AND audit_events.id > $3::uuid))"
    : "";

  const limitParam = hasCursor ? "$4" : "$2";

  const res = await db.query<PostgresAuditLogRow>(
    `
      SELECT *
      FROM (
        SELECT
          id,
          org_id,
          user_id,
          user_email,
          event_type,
          resource_type,
          resource_id,
          ip_address,
          user_agent,
          session_id,
          success,
          error_code,
          error_message,
          details,
          created_at
        FROM audit_log
        WHERE org_id = $1
        UNION ALL
        SELECT
          id,
          org_id,
          user_id,
          user_email,
          event_type,
          resource_type,
          resource_id,
          ip_address,
          user_agent,
          session_id,
          success,
          error_code,
          error_message,
          details,
          created_at
        FROM audit_log_archive
        WHERE org_id = $1
      ) AS audit_events
      ${cursorPredicate}
      ORDER BY audit_events.created_at ASC, audit_events.id ASC
      LIMIT ${limitParam}
    `,
    values
  );

  return res.rows.map((row) => {
    const createdAt = toDate(row.created_at);
    return {
      createdAt,
      event: auditLogRowToAuditEvent(row)
    } satisfies ExportableAuditEvent;
  });
}
