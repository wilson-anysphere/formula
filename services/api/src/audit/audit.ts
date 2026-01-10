import crypto from "node:crypto";
import type { Pool } from "pg";

export interface AuditEventInput {
  orgId?: string | null;
  userId?: string | null;
  userEmail?: string | null;
  eventType: string;
  resourceType: string;
  resourceId?: string | null;
  sessionId?: string | null;
  success: boolean;
  errorCode?: string | null;
  errorMessage?: string | null;
  details?: Record<string, unknown>;
  ipAddress?: string | null;
  userAgent?: string | null;
}

export async function writeAuditEvent(pool: Pool, input: AuditEventInput): Promise<void> {
  const id = crypto.randomUUID();
  await pool.query(
    `
      INSERT INTO audit_log (
        id, org_id, user_id, user_email, event_type, resource_type, resource_id,
        ip_address, user_agent, session_id,
        success, error_code, error_message, details
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14::jsonb)
    `,
    [
      id,
      input.orgId ?? null,
      input.userId ?? null,
      input.userEmail ?? null,
      input.eventType,
      input.resourceType,
      input.resourceId ?? null,
      input.ipAddress ?? null,
      input.userAgent ?? null,
      input.sessionId ?? null,
      input.success,
      input.errorCode ?? null,
      input.errorMessage ?? null,
      JSON.stringify(input.details ?? {})
    ]
  );
}

