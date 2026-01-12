import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager, DataTable, MemoryCacheStore, QueryEngine } from "@formula/power-query";

import { DocumentController } from "../../document/documentController.js";

import { getContextForDocument } from "../engine.ts";
import { refreshTableSignaturesFromBackend } from "../tableSignatures.ts";

function parseSignature(signature) {
  assert.equal(typeof signature, "string");
  const [workbookHash, definitionHash, versionStr] = signature.split(":");
  assert.ok(workbookHash, "expected non-empty workbook signature hash");
  assert.ok(definitionHash, "expected non-empty definition hash");
  const version = Number(versionStr);
  assert.ok(Number.isInteger(version), "expected integer version");
  return { workbookHash, definitionHash, version };
}

test("table signatures bump version when document edits touch the table rectangle", () => {
  const doc = new DocumentController();
  refreshTableSignaturesFromBackend(doc, [
    {
      name: "Table1",
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 1,
      end_col: 1,
      columns: ["A", "B"],
    },
  ], { workbookSignature: "workbook:test" });

  const context = getContextForDocument(doc);
  const initial = context.getTableSignature?.("Table1");
  const parsedInitial = parseSignature(initial);
  assert.equal(parsedInitial.version, 0);
  assert.equal(
    context.getTableSignature?.("table1"),
    initial,
    "expected getTableSignature to resolve table names case-insensitively",
  );

  // Cosmetic formatting edits should not invalidate table-source query caches.
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  assert.equal(context.getTableSignature?.("Table1"), initial);

  // Outside the rectangle -> no bump.
  doc.setCellValue("Sheet1", "C1", 1);
  assert.equal(context.getTableSignature?.("Table1"), initial);

  // Inside the rectangle -> bump once.
  doc.setCellValue("Sheet1", "A1", 2);
  const bumped = context.getTableSignature?.("Table1");
  assert.notEqual(bumped, initial);

  const parsedBumped = parseSignature(bumped);
  assert.equal(parsedBumped.workbookHash, parsedInitial.workbookHash);
  assert.equal(parsedBumped.definitionHash, parsedInitial.definitionHash);
  assert.equal(parsedBumped.version, parsedInitial.version + 1);
});

test("table signatures also invalidate when the table contains formulas and edits occur outside the rectangle", () => {
  const doc = new DocumentController();
  // Seed a formula cell inside the table so we conservatively treat it as depending on workbook state.
  doc.setCellFormula("Sheet1", "A1", "=C1");
  refreshTableSignaturesFromBackend(
    doc,
    [
      {
        name: "Table1",
        sheet_id: "Sheet1",
        start_row: 0,
        start_col: 0,
        end_row: 1,
        end_col: 1,
        columns: ["A", "B"],
      },
    ],
    { workbookSignature: "workbook:test" },
  );

  const context = getContextForDocument(doc);
  const initial = context.getTableSignature?.("Table1");

  // Change a cell outside the rectangle; because formulas might reference it, the signature bumps.
  doc.setCellValue("Sheet1", "C1", 1);
  assert.notEqual(context.getTableSignature?.("Table1"), initial);
});

test("QueryEngine cache keys incorporate table signatures", async () => {
  const doc = new DocumentController();
  refreshTableSignaturesFromBackend(doc, [
    {
      name: "Table1",
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 1,
      end_col: 1,
      columns: ["A"],
    },
  ], { workbookSignature: "workbook:test" });

  const context = getContextForDocument(doc);
  const engine = new QueryEngine({ cache: new CacheManager({ store: new MemoryCacheStore() }) });
  const query = { id: "q_table", name: "Table", source: { type: "table", table: "Table1" }, steps: [] };

  const key1 = await engine.getCacheKey(query, context, {});
  assert.ok(key1, "expected cache key to be computed");

  doc.setCellValue("Sheet1", "A1", 42);
  const key2 = await engine.getCacheKey(query, context, {});
  assert.ok(key2, "expected cache key to be computed after table edit");

  assert.notEqual(key2, key1);
});

test("table-source queries only cache when a signature is available", async () => {
  let getTableCalls = 0;
  const table = DataTable.fromGrid(
    [
      ["A"],
      [1],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const engine = new QueryEngine({
    cache: new CacheManager({ store: new MemoryCacheStore() }),
    tableAdapter: {
      getTable: async () => {
        getTableCalls += 1;
        return table;
      },
    },
  });

  const query = { id: "q_table_cache", name: "Table Cache", source: { type: "table", table: "Table1" }, steps: [] };

  await engine.executeQuery(query, {}, {});
  await engine.executeQuery(query, {}, {});
  assert.equal(getTableCalls, 2, "expected table adapter to run twice when signatures are missing");

  getTableCalls = 0;
  const ctx = { getTableSignature: () => "sig" };
  await engine.executeQuery(query, ctx, {});
  await engine.executeQuery(query, ctx, {});
  assert.equal(getTableCalls, 1, "expected second execution to reuse cached result when signature is present");
});

test("table signature definition hash changes (and version bumps) when the backend table definition changes", () => {
  const doc = new DocumentController();

  refreshTableSignaturesFromBackend(
    doc,
    [
      {
        name: "Table1",
        sheet_id: "Sheet1",
        start_row: 0,
        start_col: 0,
        end_row: 1,
        end_col: 0,
        columns: ["A"],
      },
    ],
    { workbookSignature: "workbook:test" },
  );

  const context = getContextForDocument(doc);
  const sig1 = context.getTableSignature?.("Table1");
  const parsed1 = parseSignature(sig1);
  assert.equal(parsed1.version, 0);

  // Simulate a resize / header change coming from the backend (e.g. after reload).
  refreshTableSignaturesFromBackend(
    doc,
    [
      {
        name: "Table1",
        sheet_id: "Sheet1",
        start_row: 0,
        start_col: 0,
        end_row: 2,
        end_col: 1,
        columns: ["A", "B"],
      },
    ],
    { workbookSignature: "workbook:test" },
  );

  const sig2 = context.getTableSignature?.("Table1");
  const parsed2 = parseSignature(sig2);
  assert.equal(parsed2.workbookHash, parsed1.workbookHash);
  assert.notEqual(parsed2.definitionHash, parsed1.definitionHash);
  assert.equal(parsed2.version, parsed1.version + 1);
});

test("table signatures are scoped by workbook signature to avoid cross-workbook cache collisions", async () => {
  const engine = new QueryEngine({ cache: new CacheManager({ store: new MemoryCacheStore() }) });
  const query = { id: "q_table", name: "Table", source: { type: "table", table: "Table1" }, steps: [] };

  const docA = new DocumentController();
  refreshTableSignaturesFromBackend(
    docA,
    [
      {
        name: "Table1",
        sheet_id: "Sheet1",
        start_row: 0,
        start_col: 0,
        end_row: 0,
        end_col: 0,
        columns: ["A"],
      },
    ],
    { workbookSignature: "workbook:A" },
  );
  const keyA = await engine.getCacheKey(query, getContextForDocument(docA), {});

  const docB = new DocumentController();
  refreshTableSignaturesFromBackend(
    docB,
    [
      {
        name: "Table1",
        sheet_id: "Sheet1",
        start_row: 0,
        start_col: 0,
        end_row: 0,
        end_col: 0,
        columns: ["A"],
      },
    ],
    { workbookSignature: "workbook:B" },
  );
  const keyB = await engine.getCacheKey(query, getContextForDocument(docB), {});

  assert.ok(keyA);
  assert.ok(keyB);
  assert.notEqual(keyB, keyA);
});
