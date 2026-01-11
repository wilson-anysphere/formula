import assert from "node:assert/strict";
import test from "node:test";
 
import { DocumentController } from "../../document/documentController.js";

import { createDesktopQueryEngine, getContextForDocument } from "../engine.ts";
import { refreshDefinedNameSignaturesFromBackend } from "../tableSignatures.ts";
 
function makeCell(value) {
  return { value, formula: null, display_value: value == null ? "" : String(value) };
}
 
test("Excel.CurrentWorkbook table source resolves defined names via get_range", async () => {
  const originalTauri = globalThis.__TAURI__;
  
  let version = 1;

  const doc = new DocumentController();
  refreshDefinedNameSignaturesFromBackend(doc, [
    { name: "MyNamedRange", refers_to: "Sheet1!$A$1:$B$3", sheet_id: null },
  ]);
  const context = getContextForDocument(doc);

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        if (cmd === "list_tables") return [];
        if (cmd === "list_defined_names") {
          return [{ name: "MyNamedRange", refers_to: "Sheet1!$A$1:$B$3", sheet_id: null }];
        }
        if (cmd === "get_range") {
          // Ensure the adapter parsed the sheet + bounds correctly.
          assert.equal(args.sheet_id, "Sheet1");
          assert.deepEqual(
            { start_row: args.start_row, start_col: args.start_col, end_row: args.end_row, end_col: args.end_col },
            { start_row: 0, start_col: 0, end_row: 2, end_col: 1 },
          );
 
          const score = version === 1 ? 10 : 15;
          return {
            start_row: 0,
            start_col: 0,
            values: [
              [makeCell("Name"), makeCell("Score")],
              [makeCell("Alice"), makeCell(score)],
              [makeCell("Bob"), makeCell(20)],
            ],
          };
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };
 
  try {
    const engine = createDesktopQueryEngine();
    const query = {
      id: "q_named_range",
      name: "NamedRange",
      source: { type: "table", table: "MyNamedRange" },
      steps: [],
    };
  
    const keyV1 = await engine.getCacheKey(query, context, {});
    assert.ok(keyV1);
  
    const first = await engine.executeQueryWithMeta(query, context, {});
    assert.equal(first.meta.cache?.hit, false);
    assert.deepEqual(first.table.toGrid(), [
      ["Name", "Score"],
      ["Alice", 10],
      ["Bob", 20],
    ]);
  
    const second = await engine.executeQueryWithMeta(query, context, {});
    assert.equal(second.meta.cache?.hit, true, "expected cache hit when named range values are unchanged");
  
    version = 2;
    doc.setCellValue("Sheet1", "B2", 15);
    const keyV2 = await engine.getCacheKey(query, context, {});
    assert.ok(keyV2);
    assert.notEqual(keyV1, keyV2, "expected cache key to change when a cell within the named range changes");
  
    const third = await engine.executeQueryWithMeta(query, context, {});
    assert.equal(third.meta.cache?.hit, false, "expected cache miss after named range edit");
    assert.deepEqual(third.table.toGrid(), [
      ["Name", "Score"],
      ["Alice", 15],
      ["Bob", 20],
    ]);
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});
 
test("Excel.CurrentWorkbook table source rejects non-range defined name formulas", async () => {
  const originalTauri = globalThis.__TAURI__;
  
  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd) => {
        if (cmd === "list_tables") return [];
        if (cmd === "list_defined_names") {
          return [{ name: "BadRange", refers_to: "=OFFSET(Sheet1!$A$1,0,0,2,2)", sheet_id: "Sheet1" }];
        }
        if (cmd === "get_range") {
          throw new Error("get_range should not be called for unsupported formulas");
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };
 
  try {
    const engine = createDesktopQueryEngine();
    const query = {
      id: "q_bad_range",
      name: "BadRange",
      source: { type: "table", table: "BadRange" },
      steps: [],
    };
  
    await assert.rejects(engine.executeQueryWithMeta(query, {}, {}), /Only simple A1 ranges are supported/);
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});
