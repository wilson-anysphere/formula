import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { runMigrations } from "../db/migrations";
import { runRetentionSweep } from "../retention";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

describe("Retention sweep (integration)", () => {
  let db: Pool;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });
  });

  afterAll(async () => {
    await db.end();
  });

  it("archives audit logs, deletes versions, purges deleted docs, and respects legal holds", async () => {
    const now = new Date("2026-01-10T00:00:00.000Z");

    const userId = crypto.randomUUID();
    const orgId = crypto.randomUUID();

    await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [
      userId,
      "admin@example.com",
      "Admin"
    ]);
    await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
    await db.query(
      `
        INSERT INTO org_settings (
          org_id,
          audit_log_retention_days,
          document_version_retention_days,
          deleted_document_retention_days,
          legal_hold_overrides_retention
        )
        VALUES ($1, 10, 10, 30, true)
      `,
      [orgId]
    );

    const dayMs = 24 * 60 * 60 * 1000;
    const twentyDaysAgo = new Date(now.getTime() - 20 * dayMs);
    const fortyDaysAgo = new Date(now.getTime() - 40 * dayMs);
    const fiveDaysAgo = new Date(now.getTime() - 5 * dayMs);

    const docActive = crypto.randomUUID();
    const docHeld = crypto.randomUUID();
    const docDeleted = crypto.randomUUID();
    const docDeletedHeld = crypto.randomUUID();

    await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1,$2,$3,$4)", [
      docActive,
      orgId,
      "Active",
      userId
    ]);
    await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1,$2,$3,$4)", [
      docHeld,
      orgId,
      "Held",
      userId
    ]);
    await db.query(
      "INSERT INTO documents (id, org_id, title, created_by, deleted_at) VALUES ($1,$2,$3,$4,$5)",
      [docDeleted, orgId, "Deleted", userId, fortyDaysAgo]
    );
    await db.query(
      "INSERT INTO documents (id, org_id, title, created_by, deleted_at) VALUES ($1,$2,$3,$4,$5)",
      [docDeletedHeld, orgId, "DeletedHeld", userId, fortyDaysAgo]
    );

    // Legal holds for held docs
    await db.query(
      "INSERT INTO document_legal_holds (document_id, org_id, enabled, created_by) VALUES ($1,$2,true,$3)",
      [docHeld, orgId, userId]
    );
    await db.query(
      "INSERT INTO document_legal_holds (document_id, org_id, enabled, created_by) VALUES ($1,$2,true,$3)",
      [docDeletedHeld, orgId, userId]
    );

    // Versions (one old version on each doc; one recent version on active doc)
    const vOldActive = crypto.randomUUID();
    const vNewActive = crypto.randomUUID();
    const vOldHeld = crypto.randomUUID();
    await db.query(
      "INSERT INTO document_versions (id, document_id, created_at, data) VALUES ($1,$2,$3,$4)",
      [vOldActive, docActive, twentyDaysAgo, Buffer.from("old")]
    );
    await db.query(
      "INSERT INTO document_versions (id, document_id, created_at, data) VALUES ($1,$2,$3,$4)",
      [vNewActive, docActive, fiveDaysAgo, Buffer.from("new")]
    );
    await db.query(
      "INSERT INTO document_versions (id, document_id, created_at, data) VALUES ($1,$2,$3,$4)",
      [vOldHeld, docHeld, twentyDaysAgo, Buffer.from("held")]
    );

    // Audit logs (one old for active doc, one old for held doc, one recent)
    const aOldActive = crypto.randomUUID();
    const aOldHeld = crypto.randomUUID();
    const aNew = crypto.randomUUID();
    await db.query(
      `
        INSERT INTO audit_log (id, org_id, event_type, resource_type, resource_id, success, details, created_at)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      `,
      [aOldActive, orgId, "document.opened", "document", docActive, true, {}, twentyDaysAgo]
    );
    await db.query(
      `
        INSERT INTO audit_log (id, org_id, event_type, resource_type, resource_id, success, details, created_at)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      `,
      [aOldHeld, orgId, "document.opened", "document", docHeld, true, {}, twentyDaysAgo]
    );
    await db.query(
      `
        INSERT INTO audit_log (id, org_id, event_type, resource_type, resource_id, success, details, created_at)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      `,
      [aNew, orgId, "document.opened", "document", docActive, true, {}, fiveDaysAgo]
    );

    const result = await runRetentionSweep(db, { now });
    expect(result).toEqual({
      auditEventsArchived: 1,
      documentVersionsDeleted: 1,
      documentsPurged: 1
    });

    const remainingVersions = await db.query(
      "SELECT id, document_id FROM document_versions ORDER BY created_at ASC"
    );
    expect(remainingVersions.rows).toEqual([
      { id: vOldHeld, document_id: docHeld },
      { id: vNewActive, document_id: docActive }
    ]);

    const remainingDocs = await db.query("SELECT id FROM documents");
    expect(remainingDocs.rows.map((r) => r.id).sort()).toEqual(
      [docActive, docDeletedHeld, docHeld].sort()
    );

    const hotAudit = await db.query("SELECT id FROM audit_log ORDER BY created_at ASC");
    expect(hotAudit.rows.map((r) => r.id)).toEqual([aOldHeld, aNew]);

    const archivedAudit = await db.query("SELECT id FROM audit_log_archive ORDER BY created_at ASC");
    expect(archivedAudit.rows.map((r) => r.id)).toEqual([aOldActive]);

    // Release holds and sweep again: remaining old items should be processed.
    await db.query(
      "UPDATE document_legal_holds SET enabled = false, released_at = now(), released_by = $2 WHERE document_id = $1",
      [docHeld, userId]
    );
    await db.query(
      "UPDATE document_legal_holds SET enabled = false, released_at = now(), released_by = $2 WHERE document_id = $1",
      [docDeletedHeld, userId]
    );

    const second = await runRetentionSweep(db, { now });
    expect(second).toEqual({
      auditEventsArchived: 1,
      documentVersionsDeleted: 1,
      documentsPurged: 1
    });

    const versionsAfterSecond = await db.query("SELECT id FROM document_versions");
    expect(versionsAfterSecond.rows.map((r) => r.id).sort()).toEqual([vNewActive]);

    const docsAfterSecond = await db.query("SELECT id FROM documents");
    expect(docsAfterSecond.rows.map((r) => r.id).sort()).toEqual([docActive, docHeld].sort());

    const hotAuditAfterSecond = await db.query("SELECT id FROM audit_log");
    expect(hotAuditAfterSecond.rows.map((r) => r.id).sort()).toEqual([aNew]);

    const archivedAfterSecond = await db.query("SELECT id FROM audit_log_archive");
    expect(archivedAfterSecond.rows.map((r) => r.id).sort()).toEqual([aOldActive, aOldHeld].sort());
  });
});
