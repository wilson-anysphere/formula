import type { Pool } from "pg";
import type { ExportableAuditEvent } from "./types";

export type AuditCursor = {
  lastCreatedAt: Date | null;
  lastEventId: string | null;
};

function parseDetails(raw: unknown): Record<string, unknown> {
  if (!raw) return {};
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") return parsed as Record<string, unknown>;
    } catch {
      return {};
    }
  }
  if (raw && typeof raw === "object") return raw as Record<string, unknown>;
  return {};
}

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

  const res = await db.query(
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
      id: String(row.id),
      timestamp: createdAt.toISOString(),
      createdAt,
      orgId: row.org_id ? String(row.org_id) : null,
      userId: row.user_id ? String(row.user_id) : null,
      userEmail: row.user_email ? String(row.user_email) : null,
      eventType: String(row.event_type),
      resourceType: String(row.resource_type),
      resourceId: row.resource_id ? String(row.resource_id) : null,
      ipAddress: row.ip_address ? String(row.ip_address) : null,
      userAgent: row.user_agent ? String(row.user_agent) : null,
      sessionId: row.session_id ? String(row.session_id) : null,
      success: Boolean(row.success),
      errorCode: row.error_code ? String(row.error_code) : null,
      errorMessage: row.error_message ? String(row.error_message) : null,
      details: parseDetails(row.details)
    } satisfies ExportableAuditEvent;
  });
}
