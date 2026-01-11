import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../../../../packages/power-query/src/cache/cache.js";
import { MemoryCacheStore } from "../../../../../packages/power-query/src/cache/memory.js";
import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";

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
