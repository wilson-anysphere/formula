import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { runMigrations } from "../db/migrations";
import { DLP_ACTION } from "../dlp/dlp";
import { evaluateDocumentDlpPolicy } from "../dlp/effective";
import { normalizeSelectorColumns, selectorKey } from "../dlp/classificationResolver";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

describe("DLP policy merging (org + document)", () => {
  let db: Pool;
  let orgId: string;
  let docId: string;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    const userId = crypto.randomUUID();
    orgId = crypto.randomUUID();
    docId = crypto.randomUUID();

    await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [
      userId,
      "dlp-policy-merge@example.com",
      "DLP Policy Merge"
    ]);
    await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "DLP Policy Merge Org"]);
    await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1, $2, $3, $4)", [
      docId,
      orgId,
      "DLP Policy Merge Doc",
      userId
    ]);
  });

  afterAll(async () => {
    await db.end();
  });

  beforeEach(async () => {
    await db.query("DELETE FROM document_classifications");
    await db.query("DELETE FROM document_dlp_policies");
    await db.query("DELETE FROM org_dlp_policies");
  });

  async function insertDocumentClassification(classification: any): Promise<void> {
    const selector = { scope: "document", documentId: docId };
    const cols = normalizeSelectorColumns(selector);
    const key = selectorKey(selector);

    await db.query(
      `
        INSERT INTO document_classifications (
          id,
          document_id,
          selector_key,
          selector,
          classification,
          scope,
          sheet_id,
          table_id,
          row,
          col,
          start_row,
          start_col,
          end_row,
          end_col,
          column_index,
          column_id
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16)
      `,
      [
        crypto.randomUUID(),
        docId,
        key,
        JSON.stringify(selector),
        JSON.stringify(classification),
        cols.scope,
        cols.sheetId,
        cols.tableId,
        cols.row,
        cols.col,
        cols.startRow,
        cols.startCol,
        cols.endRow,
        cols.endCol,
        cols.columnIndex,
        cols.columnId
      ]
    );
  }

  it("does not allow a document override to loosen org AI cloud processing controls", async () => {
    await db.query("INSERT INTO org_dlp_policies (org_id, policy) VALUES ($1, $2)", [
      orgId,
      JSON.stringify({
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: false
          }
        }
      })
    ]);

    // Doc policy attempts to loosen thresholds and enable Restricted inclusion + redaction.
    await db.query("INSERT INTO document_dlp_policies (document_id, policy) VALUES ($1, $2)", [
      docId,
      JSON.stringify({
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Restricted",
            allowRestrictedContent: true,
            redactDisallowed: true
          }
        }
      })
    ]);

    await insertDocumentClassification({ level: "Restricted", labels: ["PII"] });

    const evaluationNoInclude = await evaluateDocumentDlpPolicy(db, {
      orgId,
      docId,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      options: { includeRestrictedContent: false }
    });
    // Org policy forbids redaction, so the request should be blocked (not redacted),
    // even though the document policy tries to enable redaction and loosen thresholds.
    expect(evaluationNoInclude).toMatchObject({
      decision: "block",
      maxAllowed: "Confidential",
      classification: { level: "Restricted", labels: ["PII"] }
    });

    const evaluationInclude = await evaluateDocumentDlpPolicy(db, {
      orgId,
      docId,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      options: { includeRestrictedContent: true }
    });
    // Explicit inclusion must still be blocked since org disallows allowRestrictedContent.
    expect(evaluationInclude).toMatchObject({
      decision: "block",
      maxAllowed: "Confidential",
      classification: { level: "Restricted", labels: ["PII"] }
    });
  });
});

