import type { Pool } from "pg";
import { withTransaction } from "./db/tx";

export interface RetentionSweepResult {
  auditEventsArchived: number;
  documentVersionsDeleted: number;
  documentsPurged: number;
  /**
   * Number of failures from `onDocumentPurged` (for example, failed sync-server
   * state purges). Present only when `onDocumentPurged` is provided.
   */
  syncPurgesFailed?: number;
}

export async function runRetentionSweep(
  pool: Pool,
  {
    now = new Date(),
    onDocumentPurged
  }: {
    now?: Date;
    onDocumentPurged?: (args: { orgId: string; docId: string }) => Promise<void>;
  } = {}
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
  let syncPurgesFailed = 0;

  for (const org of orgs.rows) {
    const auditCutoff = new Date(now.getTime() - org.audit_log_retention_days * 24 * 60 * 60 * 1000);
    const versionCutoff = new Date(now.getTime() - org.document_version_retention_days * 24 * 60 * 60 * 1000);
    const deletedDocCutoff = new Date(
      now.getTime() - org.deleted_document_retention_days * 24 * 60 * 60 * 1000
    );
    const legalHoldOverridesRetention = org.legal_hold_overrides_retention;

    const { archivedCount, versionsDeletedCount, docsPurgedRes } =
      await withTransaction(pool, async (client) => {
        // Move audit events to cold storage before deleting from the hot audit_log table.
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

        const docsPurgedRes = await client.query<{ id: string }>(
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
            RETURNING id
          `,
          [org.org_id, deletedDocCutoff, legalHoldOverridesRetention]
        );

        return {
          archivedCount: archived.rowCount ?? 0,
          versionsDeletedCount: versionsDeleted.rowCount ?? 0,
          docsPurgedRes
        };
      });

    auditEventsArchived += archivedCount;
    documentVersionsDeleted += versionsDeletedCount;
    documentsPurged += docsPurgedRes.rowCount ?? 0;

    if (onDocumentPurged) {
      for (const row of docsPurgedRes.rows) {
        try {
          await onDocumentPurged({ orgId: org.org_id, docId: row.id });
        } catch {
          syncPurgesFailed += 1;
        }
      }
    }
  }

  const result: RetentionSweepResult = {
    auditEventsArchived,
    documentVersionsDeleted,
    documentsPurged
  };
  if (onDocumentPurged) result.syncPurgesFailed = syncPurgesFailed;
  return result;
}
