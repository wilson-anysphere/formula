import type { Pool } from "pg";
import { withTransaction } from "./db/tx";

export interface RetentionSweepResult {
  auditEventsArchived: number;
  documentVersionsDeleted: number;
  documentsPurged: number;
}

export async function runRetentionSweep(
  pool: Pool,
  { now = new Date() }: { now?: Date } = {}
): Promise<RetentionSweepResult> {
  const orgs = await pool.query<{
    org_id: string;
    audit_log_retention_days: number;
    document_version_retention_days: number;
    deleted_document_retention_days: number;
    legal_hold_overrides_retention: boolean;
  }>(
    `
      SELECT org_id,
             audit_log_retention_days,
             document_version_retention_days,
             deleted_document_retention_days,
             legal_hold_overrides_retention
      FROM org_settings
    `
  );

  let auditEventsArchived = 0;
  let documentVersionsDeleted = 0;
  let documentsPurged = 0;

  for (const org of orgs.rows) {
    await withTransaction(pool, async (client) => {
      const auditCutoff = new Date(now.getTime() - org.audit_log_retention_days * 24 * 60 * 60 * 1000);
      const versionCutoff = new Date(
        now.getTime() - org.document_version_retention_days * 24 * 60 * 60 * 1000
      );
      const deletedDocCutoff = new Date(
        now.getTime() - org.deleted_document_retention_days * 24 * 60 * 60 * 1000
      );
      const legalHoldOverridesRetention = org.legal_hold_overrides_retention;

      // Move audit events to cold storage before deleting from the hot audit_log table.
      //
      // Note: In Postgres we could do this atomically with `WITH moved AS (DELETE ... RETURNING ...) INSERT ...`.
      // pg-mem (used for tests) does not support that pattern, so we do it in a transaction.
      const archived = await client.query(
        `
          INSERT INTO audit_log_archive (
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
            created_at,
            archived_at
          )
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
            created_at,
            now()
          FROM audit_log
          WHERE org_id = $1
            AND created_at < $2
            AND (
              $3::boolean = false
              OR resource_type <> 'document'
              OR resource_id IS NULL
              OR resource_id NOT IN (
                SELECT h.document_id::text
                FROM document_legal_holds h
                WHERE h.org_id = $1
                  AND h.enabled = true
              )
            )
          RETURNING id
        `,
        [org.org_id, auditCutoff, legalHoldOverridesRetention]
      );
      auditEventsArchived += archived.rowCount ?? 0;

      await client.query(
        `
          DELETE FROM audit_log
          WHERE org_id = $1
            AND created_at < $2
            AND (
              $3::boolean = false
              OR resource_type <> 'document'
              OR resource_id IS NULL
              OR resource_id NOT IN (
                SELECT h.document_id::text
                FROM document_legal_holds h
                WHERE h.org_id = $1
                  AND h.enabled = true
              )
            )
        `,
        [org.org_id, auditCutoff, legalHoldOverridesRetention]
      );

      const versionsDeleted = await client.query(
        `
          DELETE FROM document_versions
          WHERE document_id IN (
            SELECT d.id
            FROM documents d
            WHERE d.org_id = $1
          )
            AND created_at < $2
            AND (
              $3::boolean = false
              OR document_id NOT IN (
                SELECT h.document_id
                FROM document_legal_holds h
                WHERE h.org_id = $1
                  AND h.enabled = true
              )
            )
        `,
        [org.org_id, versionCutoff, legalHoldOverridesRetention]
      );
      documentVersionsDeleted += versionsDeleted.rowCount ?? 0;

      const docsPurgedRes = await client.query(
        `
          DELETE FROM documents
          WHERE org_id = $1
            AND deleted_at IS NOT NULL
            AND deleted_at < $2
            AND (
              $3::boolean = false
              OR id NOT IN (
                SELECT h.document_id
                FROM document_legal_holds h
                WHERE h.org_id = $1
                  AND h.enabled = true
              )
            )
        `,
        [org.org_id, deletedDocCutoff, legalHoldOverridesRetention]
      );
      documentsPurged += docsPurgedRes.rowCount ?? 0;
    });
  }

  return { auditEventsArchived, documentVersionsDeleted, documentsPurged };
}
