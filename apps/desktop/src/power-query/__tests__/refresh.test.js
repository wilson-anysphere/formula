import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../../../../packages/power-query/src/table.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryRefreshManager } from "../refresh.ts";

function makeMeta(queryId, table) {
  return {
    queryId,
    startedAt: new Date(0),
    completedAt: new Date(0),
    refreshedAt: new Date(0),
    sources: [],
    outputSchema: { columns: table.columns, inferred: true },
    outputRowCount: table.rowCount,
  };
}

class StaticEngine {
  /**
   * @param {DataTable} table
   */
  constructor(table) {
    this.table = table;
  }

  async executeQueryWithMeta(query, _context, options) {
    options?.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey: "k" });
    if (options?.signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }
    return { table: this.table, meta: makeMeta(query.id, this.table) };
  }
}

test("DesktopPowerQueryRefreshManager applies completed refresh results into the destination", async () => {
  const table = DataTable.fromGrid(
    [
      ["A", "B"],
      [1, 2],
      [3, 4],
    ],
    { hasHeaders: true, inferTypes: true },
  );
  const engine = new StaticEngine(table);
  const doc = new DocumentController({ engine: new MockEngine() });

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1 });

  const query = {
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  mgr.registerQuery(query);

  const applied = new Promise((resolve) => {
    const unsub = mgr.onEvent((evt) => {
      if (evt.type === "apply:completed" && evt.queryId === query.id) {
        unsub();
        resolve(evt);
      }
    });
  });

  const handle = mgr.refresh(query.id);
  await handle.promise;
  await applied;

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "A");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, 2);
  assert.equal(doc.getCell("Sheet1", { row: 2, col: 0 }).value, 3);

  mgr.dispose();
});

test("DesktopPowerQueryRefreshManager cancellation aborts apply and reverts partial writes", async () => {
  const table = DataTable.fromGrid(
    [
      ["A"],
      [1],
      [2],
      [3],
    ],
    { hasHeaders: true, inferTypes: true },
  );
  const engine = new StaticEngine(table);
  const doc = new DocumentController({ engine: new MockEngine() });

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1 });

  const query = {
    id: "q_cancel",
    name: "Cancel",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };
  mgr.registerQuery(query);

  const handle = mgr.refresh(query.id);

  let cancelled = false;
  const done = new Promise((resolve) => {
    const unsub = mgr.onEvent((evt) => {
      if (evt.type === "apply:progress" && evt.queryId === query.id && !cancelled) {
        cancelled = true;
        handle.cancel();
      }
      if (evt.type === "apply:cancelled" && evt.queryId === query.id) {
        unsub();
        resolve(evt);
      }
    });
  });

  await handle.promise;
  await done;

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, null);
  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.equal(doc.batchDepth, 0);

  mgr.dispose();
});

