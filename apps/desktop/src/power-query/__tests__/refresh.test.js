import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../../../../packages/power-query/src/table.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryRefreshManager } from "../refresh.ts";
import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";

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

class OrderedEngine {
  constructor(resultsById) {
    this.resultsById = resultsById;
    this.calls = [];
  }

  createSession() {
    return { credentialCache: new Map(), permissionCache: new Map() };
  }

  async executeQueryWithMetaInSession(query, _context, options) {
    return this.executeQueryWithMeta(query, _context, options);
  }

  async executeQueryWithMeta(query, _context, options) {
    this.calls.push(query.id);
    options?.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey: "k" });
    if (options?.signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      throw err;
    }
    const table = this.resultsById[query.id];
    if (!table) throw new Error(`Missing table for ${query.id}`);
    return { table, meta: makeMeta(query.id, table) };
  }
}

function deferred() {
  /** @type {(value: any) => void} */
  let resolve;
  /** @type {(reason?: any) => void} */
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

class ControlledEngine {
  constructor() {
    /** @type {{ queryId: string, deferred: ReturnType<typeof deferred>, signal?: AbortSignal }[]} */
    this.calls = [];
  }

  createSession() {
    return { credentialCache: new Map(), permissionCache: new Map() };
  }

  executeQueryWithMetaInSession(query, context, options) {
    return this.executeQueryWithMeta(query, context, options);
  }

  executeQueryWithMeta(query, _context, options) {
    const d = deferred();
    this.calls.push({ queryId: query.id, deferred: d, signal: options?.signal });
    options?.signal?.addEventListener("abort", () => {
      const err = new Error("Aborted");
      err.name = "AbortError";
      d.reject(err);
    });
    return d.promise;
  }
}

test("DesktopPowerQueryRefreshManager persists refresh state when a stateStore is provided", async () => {
  const table = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const engine = new StaticEngine(table);
  const doc = new DocumentController({ engine: new MockEngine() });

  let loadCalls = 0;
  /** @type {any} */
  let savedState = null;

  const stateStore = {
    load: async () => {
      loadCalls += 1;
      return {};
    },
    save: async (state) => {
      savedState = state;
    },
  };

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1, stateStore });

  const query = {
    id: "q_state",
    name: "State",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "interval", intervalMs: 10_000 },
  };

  mgr.registerQuery(query);
  await mgr.manager.ready;
  await new Promise((resolve) => queueMicrotask(resolve));

  assert.equal(loadCalls, 1);
  assert.ok(savedState);
  assert.ok(savedState[query.id]);

  mgr.dispose();
});

test("DesktopPowerQueryRefreshManager refreshAll respects dependencies and applies refreshed outputs", async () => {
  const tableA = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const tableB = DataTable.fromGrid([["B"], [2]], { hasHeaders: true, inferTypes: true });
  const engine = new OrderedEngine({ A: tableA, B: tableB });
  const doc = new DocumentController({ engine: new MockEngine() });

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 2, batchSize: 1 });

  const qA = {
    id: "A",
    name: "A",
    source: { type: "range", range: { values: [["x"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  const qB = {
    id: "B",
    name: "B",
    source: { type: "query", queryId: "A" },
    steps: [],
    destination: { sheetId: "Sheet2", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  mgr.registerQuery(qA);
  mgr.registerQuery(qB);

  const applied = new Set();
  const appliedPromise = new Promise((resolve, reject) => {
    const unsub = mgr.onEvent((evt) => {
      if (evt.type === "apply:error") {
        unsub();
        reject(evt.error);
      }
      if (evt.type === "apply:completed") {
        applied.add(evt.queryId);
        if (applied.has("A") && applied.has("B")) {
          unsub();
          resolve(undefined);
        }
      }
    });
  });

  const handle = mgr.refreshAll(["B"]);
  await handle.promise;
  await appliedPromise;

  assert.deepEqual(engine.calls, ["A", "B"]);

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "A");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, 1);
  assert.equal(doc.getCell("Sheet2", { row: 0, col: 0 }).value, "B");
  assert.equal(doc.getCell("Sheet2", { row: 1, col: 0 }).value, 2);

  mgr.dispose();
});

test("DesktopPowerQueryRefreshManager refreshAll shares credential prompts across the session", async () => {
  let credentialRequests = 0;
  const engine = new QueryEngine({
    fileAdapter: { readText: async () => "Value\n1\n" },
    onCredentialRequest: async () => {
      credentialRequests += 1;
      return { token: "ok" };
    },
  });

  const doc = new DocumentController({ engine: new MockEngine() });
  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 2, batchSize: 1 });

  mgr.registerQuery({ id: "Q1", name: "Q1", source: { type: "csv", path: "file.csv" }, steps: [], refreshPolicy: { type: "manual" } });
  mgr.registerQuery({ id: "Q2", name: "Q2", source: { type: "csv", path: "file.csv" }, steps: [], refreshPolicy: { type: "manual" } });

  await mgr.refreshAll(["Q1", "Q2"]).promise;
  assert.equal(credentialRequests, 1);

  mgr.dispose();
});

test("DesktopPowerQueryRefreshManager dispose cancels in-flight refreshAll sessions", async () => {
  const engine = new ControlledEngine();
  const doc = new DocumentController({ engine: new MockEngine() });
  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1 });

  mgr.registerQuery({ id: "q1", name: "Q1", source: { type: "range", range: { values: [["x"], [1]], hasHeaders: true } }, steps: [] });
  mgr.registerQuery({ id: "q2", name: "Q2", source: { type: "range", range: { values: [["x"], [2]], hasHeaders: true } }, steps: [] });

  const handle = mgr.refreshAll(["q1", "q2"]);
  assert.equal(engine.calls.length, 1, "only one job should start with concurrency=1");

  mgr.dispose();
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  assert.equal(engine.calls.length, 1, "dispose should cancel queued work before it starts");
});

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
