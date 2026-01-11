import type { Pool } from "pg";

import {
  assertAuditEvent,
  buildPostgresAuditLogInsert,
  createAuditEvent,
  type AuditEvent
} from "@formula/audit-core";

export { createAuditEvent };
export type { AuditEvent };

export async function writeAuditEvent(pool: Pool, event: AuditEvent): Promise<void> {
  assertAuditEvent(event);
  const { text, values } = buildPostgresAuditLogInsert(event);
  await pool.query(text, values);
}
