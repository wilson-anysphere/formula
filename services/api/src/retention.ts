import type { Pool } from "pg";

export interface RetentionSweepResult {
  auditEventsDeleted: number;
  documentVersionsDeleted: number;
}

export async function runRetentionSweep(pool: Pool): Promise<RetentionSweepResult> {
  const orgs = await pool.query<{
    org_id: string;
    audit_log_retention_days: number;
    document_version_retention_days: number;
  }>(
    `
      SELECT org_id, audit_log_retention_days, document_version_retention_days
      FROM org_settings
    `
  );

  let auditEventsDeleted = 0;
  let documentVersionsDeleted = 0;

  for (const org of orgs.rows) {
    const auditCutoff = new Date(Date.now() - org.audit_log_retention_days * 24 * 60 * 60 * 1000);
    const versionCutoff = new Date(
      Date.now() - org.document_version_retention_days * 24 * 60 * 60 * 1000
    );

    const auditDeleted = await pool.query(
      "DELETE FROM audit_log WHERE org_id = $1 AND created_at < $2",
      [org.org_id, auditCutoff]
    );
    auditEventsDeleted += auditDeleted.rowCount ?? 0;

    const versionsDeleted = await pool.query(
      `
        DELETE FROM document_versions v
        USING documents d
        WHERE v.document_id = d.id
          AND d.org_id = $1
          AND v.created_at < $2
      `,
      [org.org_id, versionCutoff]
    );
    documentVersionsDeleted += versionsDeleted.rowCount ?? 0;
  }

  return { auditEventsDeleted, documentVersionsDeleted };
}

