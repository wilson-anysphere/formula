import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { runMigrations } from "../db/migrations";
import {
  getAggregateClassificationForRange,
  getEffectiveClassificationForSelector,
  normalizeSelectorColumns,
  selectorKey
} from "../dlp/classificationResolver";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

describe("classificationResolver: selector precedence", () => {
  let db: Pool;
  let docId: string;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    const userId = crypto.randomUUID();
    const orgId = crypto.randomUUID();
    docId = crypto.randomUUID();

    await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [
      userId,
      "resolver@example.com",
      "Resolver"
    ]);
    await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Resolver Org"]);
    await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1, $2, $3, $4)", [
      docId,
      orgId,
      "Resolver Doc",
      userId
    ]);
  });

  afterAll(async () => {
    await db.end();
  });

  beforeEach(async () => {
    await db.query("DELETE FROM document_classifications");
  });

  async function insertClassification(selector: any, classification: any): Promise<string> {
    const key = selectorKey(selector);
    const cols = normalizeSelectorColumns(selector);

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

    return key;
  }

  it("cell overrides range", async () => {
    const rangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
    };
    await insertClassification(rangeSelector, { level: "Restricted", labels: ["PII"] });

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 1,
      col: 1
    };
    const cellKey = await insertClassification(cellSelector, { level: "Internal", labels: ["OK"] });

    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toMatchObject({
      classification: { level: "Internal", labels: ["OK"] },
      source: { scope: "cell", selectorKey: cellKey }
    });
  });

  it("smallest containing range wins", async () => {
    const bigRangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 3, col: 3 } }
    };
    await insertClassification(bigRangeSelector, { level: "Restricted", labels: ["PII"] });

    const smallRangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } }
    };
    const smallKey = await insertClassification(smallRangeSelector, { level: "Internal", labels: ["Subset"] });

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 1,
      col: 1
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toMatchObject({
      classification: { level: "Internal", labels: ["Subset"] },
      source: { scope: "range", selectorKey: smallKey }
    });
  });

  it("overlapping ranges with same specificity pick max classification", async () => {
    const rangeA = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }
    };
    await insertClassification(rangeA, { level: "Internal", labels: ["A"] });

    const rangeB = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 1, col: 0 }, end: { row: 2, col: 1 } }
    };
    const keyB = await insertClassification(rangeB, { level: "Confidential", labels: ["B"] });

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 1,
      col: 1
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved.classification).toEqual({ level: "Confidential", labels: ["A", "B"] });
    expect(resolved.source).toEqual({ scope: "range", selectorKey: keyB });
  });

  it("merges labels across equally-specific Restricted ranges deterministically", async () => {
    const rangeA = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }
    };
    const keyA = await insertClassification(rangeA, { level: "Restricted", labels: ["A"] });

    const rangeB = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 1, col: 0 }, end: { row: 2, col: 1 } }
    };
    const keyB = await insertClassification(rangeB, { level: "Restricted", labels: ["B"] });

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 1,
      col: 1
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved.classification).toEqual({ level: "Restricted", labels: ["A", "B"] });
    expect(resolved.source).toEqual({
      scope: "range",
      selectorKey: [keyA, keyB].sort()[0]
    });
  });

  it("aggregate classification unions labels across all intersecting selectors", async () => {
    await insertClassification({ scope: "document", documentId: docId }, { level: "Internal", labels: ["Doc"] });

    const rangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }
    };
    await insertClassification(rangeSelector, { level: "Restricted", labels: ["PII"] });

    await insertClassification(
      { scope: "cell", documentId: docId, sheetId: "Sheet1", row: 0, col: 0 },
      { level: "Confidential", labels: ["Cell"] }
    );

    const aggregate = await getAggregateClassificationForRange(db, docId, "Sheet1", 0, 0, 0, 0);
    expect(aggregate).toEqual({ level: "Restricted", labels: ["Cell", "Doc", "PII"] });
  });

  it("falls back to sheet classification when no more specific selector matches", async () => {
    const sheetSelector = { scope: "sheet", documentId: docId, sheetId: "Sheet1" };
    const sheetKey = await insertClassification(sheetSelector, { level: "Confidential", labels: ["Sheet"] });

    const cellSelector = { scope: "cell", documentId: docId, sheetId: "Sheet1", row: 0, col: 0 };
    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toEqual({
      classification: { level: "Confidential", labels: ["Sheet"] },
      source: { scope: "sheet", selectorKey: sheetKey }
    });
  });

  it("column classification overrides sheet/document even if less restrictive", async () => {
    await insertClassification({ scope: "document", documentId: docId }, { level: "Restricted", labels: ["Doc"] });
    await insertClassification(
      { scope: "sheet", documentId: docId, sheetId: "Sheet1" },
      { level: "Confidential", labels: ["Sheet"] }
    );

    const columnSelector = { scope: "column", documentId: docId, sheetId: "Sheet1", columnIndex: 0 };
    const columnKey = await insertClassification(columnSelector, { level: "Internal", labels: ["Col"] });

    const cellSelector = { scope: "cell", documentId: docId, sheetId: "Sheet1", row: 5, col: 0 };
    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toEqual({
      classification: { level: "Internal", labels: ["Col"] },
      source: { scope: "column", selectorKey: columnKey }
    });
  });

  it("range classification overrides column classification for a cell", async () => {
    await insertClassification(
      { scope: "column", documentId: docId, sheetId: "Sheet1", columnIndex: 1 },
      { level: "Restricted", labels: ["ColumnRestricted"] }
    );

    const rangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 1 }, end: { row: 2, col: 1 } }
    };
    const rangeKey = await insertClassification(rangeSelector, { level: "Internal", labels: ["RangeOverride"] });

    const cellSelector = { scope: "cell", documentId: docId, sheetId: "Sheet1", row: 1, col: 1 };
    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toEqual({
      classification: { level: "Internal", labels: ["RangeOverride"] },
      source: { scope: "range", selectorKey: rangeKey }
    });
  });

  it("table column classification (tableId + columnId) overrides sheet column classification when table context is provided", async () => {
    await insertClassification({ scope: "document", documentId: docId }, { level: "Restricted", labels: ["Doc"] });
    await insertClassification(
      { scope: "column", documentId: docId, sheetId: "Sheet1", columnIndex: 3 },
      { level: "Restricted", labels: ["SheetCol"] }
    );

    const tableColumnSelector = {
      scope: "column",
      documentId: docId,
      sheetId: "Sheet1",
      tableId: "Table1",
      columnId: "ColA"
    };
    const columnKey = await insertClassification(tableColumnSelector, { level: "Internal", labels: ["TableCol"] });

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 10,
      col: 3,
      tableId: "Table1",
      columnId: "ColA"
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, cellSelector);
    expect(resolved).toEqual({
      classification: { level: "Internal", labels: ["TableCol"] },
      source: { scope: "column", selectorKey: columnKey }
    });
  });

  it("returns Public/default when no classifications exist", async () => {
    const resolved = await getEffectiveClassificationForSelector(db, docId, {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 0,
      col: 0
    });
    expect(resolved).toEqual({
      classification: { level: "Public", labels: [] },
      source: { scope: "default", selectorKey: "default" }
    });
  });

  it("range selector falls back to column classification only when it is a single column", async () => {
    const columnSelector = { scope: "column", documentId: docId, sheetId: "Sheet1", columnIndex: 0 };
    const columnKey = await insertClassification(columnSelector, { level: "Confidential", labels: ["Col"] });

    const singleColumnRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } }
    };
    const resolvedSingle = await getEffectiveClassificationForSelector(db, docId, singleColumnRange);
    expect(resolvedSingle).toEqual({
      classification: { level: "Confidential", labels: ["Col"] },
      source: { scope: "column", selectorKey: columnKey }
    });

    await insertClassification(
      { scope: "sheet", documentId: docId, sheetId: "Sheet1" },
      { level: "Internal", labels: ["Sheet"] }
    );

    const multiColumnRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }
    };
    const resolvedMulti = await getEffectiveClassificationForSelector(db, docId, multiColumnRange);
    expect(resolvedMulti).toMatchObject({
      classification: { level: "Internal", labels: ["Sheet"] },
      source: { scope: "sheet" }
    });
  });

  it("range selector uses the smallest containing range classification", async () => {
    const bigRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 10, col: 10 } }
    };
    await insertClassification(bigRange, { level: "Restricted", labels: ["Big"] });

    const smallRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 2, col: 2 }, end: { row: 5, col: 5 } }
    };
    const smallKey = await insertClassification(smallRange, { level: "Internal", labels: ["Small"] });

    const targetRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 3, col: 3 }, end: { row: 4, col: 4 } }
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, targetRange);
    expect(resolved).toEqual({
      classification: { level: "Internal", labels: ["Small"] },
      source: { scope: "range", selectorKey: smallKey }
    });
  });

  it("range selector merges labels across equally-small containing ranges", async () => {
    const rangeA = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 4, col: 4 } }
    };
    await insertClassification(rangeA, { level: "Confidential", labels: ["A"] });

    // Same selectorKey would conflict; use a distinct but equally-sized range.
    const rangeB2 = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 1, col: 1 }, end: { row: 5, col: 5 } }
    };
    const keyB = await insertClassification(rangeB2, { level: "Restricted", labels: ["B"] });

    const targetRange = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 2, col: 2 }, end: { row: 3, col: 3 } }
    };

    const resolved = await getEffectiveClassificationForSelector(db, docId, targetRange);
    expect(resolved.classification).toEqual({ level: "Restricted", labels: ["A", "B"] });
    expect(resolved.source).toEqual({ scope: "range", selectorKey: keyB });
  });
});
