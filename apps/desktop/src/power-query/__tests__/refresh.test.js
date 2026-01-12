import assert from "node:assert/strict";
import test from "node:test";

import { DataTable, QueryEngine, RefreshManager } from "@formula/power-query";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryRefreshManager } from "../refresh.ts";
import { createPowerQueryRefreshStateStore } from "../refreshStateStore.ts";

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

test("DesktopPowerQueryRefreshManager restores persisted refresh policy when the query omits refreshPolicy", async () => {
  const table = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const engine = new StaticEngine(table);
  const doc = new DocumentController({ engine: new MockEngine() });

  /** @type {any} */
  let savedState = null;

  const stateStore = {
    load: async () => ({
      q_persisted: { policy: { type: "interval", intervalMs: 10_000 }, lastRunAtMs: 123 },
    }),
    save: async (state) => {
      savedState = state;
    },
  };

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1, stateStore });

  const query = {
    id: "q_persisted",
    name: "Persisted",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
  };

  mgr.registerQuery(query);
  await mgr.ready;
  await new Promise((resolve) => queueMicrotask(resolve));

  assert.equal(savedState?.q_persisted?.policy?.type, "interval");
  assert.equal(savedState?.q_persisted?.policy?.intervalMs, 10_000);

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

test("DesktopPowerQueryRefreshManager refreshWithDependencies refreshes dependencies and returns the target result", async () => {
  const tableA = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const tableB = DataTable.fromGrid([["B"], [2]], { hasHeaders: true, inferTypes: true });
  const engine = new OrderedEngine({ A: tableA, B: tableB });
  const doc = new DocumentController({ engine: new MockEngine() });

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 2, batchSize: 1 });

  mgr.registerQuery({ id: "A", name: "A", source: { type: "range", range: { values: [["x"], [1]], hasHeaders: true } }, steps: [] });
  mgr.registerQuery({ id: "B", name: "B", source: { type: "query", queryId: "A" }, steps: [] });

  const result = await mgr.refreshWithDependencies("B").promise;
  assert.deepEqual(engine.calls, ["A", "B"]);
  assert.equal(result.meta.queryId, "B");

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

test("DesktopPowerQueryRefreshManager refreshAll updates lastRunAtMs in the refresh stateStore", async () => {
  const table = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const engine = new StaticEngine(table);
  const doc = new DocumentController({ engine: new MockEngine() });

  /** @type {any[]} */
  const savedStates = [];
  const stateStore = {
    load: async () => ({}),
    save: async (state) => {
      savedStates.push(state);
    },
  };

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 1, batchSize: 1, stateStore });

  const query = {
    id: "q_interval",
    name: "Interval",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "interval", intervalMs: 10_000 },
  };

  mgr.registerQuery(query);
  await mgr.manager.ready;
  await new Promise((resolve) => queueMicrotask(resolve));

  const before = savedStates.at(-1)?.[query.id]?.lastRunAtMs;
  assert.equal(before, undefined);

  await mgr.refreshAll([query.id]).promise;
  await new Promise((resolve) => setImmediate(resolve));

  const afterEntry = savedStates.at(-1)?.[query.id];
  assert.ok(afterEntry);
  assert.equal(afterEntry.policy?.type, "interval");
  assert.equal(typeof afterEntry.lastRunAtMs, "number");

  mgr.dispose();
});

test("DesktopPowerQueryRefreshManager refreshAll cancelQuery aborts the apply phase for that query only", async () => {
  const bigTable = DataTable.fromGrid([["A"], ...Array.from({ length: 50 }, (_, i) => [i + 1])], { hasHeaders: true, inferTypes: true });
  const smallTable = DataTable.fromGrid([["B"], [1]], { hasHeaders: true, inferTypes: true });
  const engine = new OrderedEngine({ q1: bigTable, q2: smallTable });
  const doc = new DocumentController({ engine: new MockEngine() });

  const mgr = new DesktopPowerQueryRefreshManager({ engine, document: doc, concurrency: 2, batchSize: 1 });

  mgr.registerQuery({
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  mgr.registerQuery({
    id: "q2",
    name: "Q2",
    source: { type: "range", range: { values: [["X"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet2", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  const handle = mgr.refreshAll(["q1", "q2"]);

  let cancelled = false;
  let q1Cancelled = false;
  let q2Completed = false;

  const done = new Promise((resolve, reject) => {
    const unsub = mgr.onEvent((evt) => {
      if (evt.type === "apply:error") {
        unsub();
        reject(evt.error);
        return;
      }

      if (evt.type === "apply:progress" && evt.queryId === "q1" && !cancelled) {
        cancelled = true;
        handle.cancelQuery?.("q1");
      }

      if (evt.type === "apply:cancelled" && evt.queryId === "q1") {
        q1Cancelled = true;
      }

      if (evt.type === "apply:completed" && evt.queryId === "q2") {
        q2Completed = true;
      }

      if (q1Cancelled && q2Completed) {
        unsub();
        resolve(undefined);
      }
    });
  });

  await handle.promise;
  await done;

  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.equal(doc.getCell("Sheet2", { row: 0, col: 0 }).value, "B");

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
function makeQuery(id, refreshPolicy) {
  const query = {
    id,
    name: id,
    source: { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } },
    steps: [],
  };
  if (refreshPolicy) query.refreshPolicy = refreshPolicy;
  return query;
}

function makeResult(queryId) {
  const table = new DataTable([], []);
  return { table, meta: makeMeta(queryId, table) };
}

class FakeTimers {
  constructor(now = 0) {
    this.now = now;
    this.nextId = 1;
    /** @type {Map<number, { time: number, fn: () => void }>} */
    this.tasks = new Map();
  }

  setTimeout(fn, ms) {
    const id = this.nextId++;
    this.tasks.set(id, { time: this.now + ms, fn });
    return id;
  }

  clearTimeout(id) {
    this.tasks.delete(id);
  }

  advance(ms) {
    this.now += ms;
    while (true) {
      let next = null;
      for (const [id, task] of this.tasks) {
        if (task.time <= this.now) {
          if (!next || task.time < next.task.time) next = { id, task };
        }
      }
      if (!next) break;
      this.tasks.delete(next.id);
      next.task.fn();
    }
  }
}

class MapStorage {
  constructor() {
    this.map = new Map();
  }

  getItem(key) {
    return this.map.get(key) ?? null;
  }

  setItem(key, value) {
    this.map.set(key, value);
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

test("RefreshManager persists interval schedules via the desktop refresh state store", async () => {
  const storage = new MapStorage();
  const stateStore1 = createPowerQueryRefreshStateStore({ workbookId: "wb_refresh", storage });
  const timers = new FakeTimers();

  const engine1 = new ControlledEngine();
  const manager1 = new RefreshManager({
    engine: engine1,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    stateStore: stateStore1,
  });

  manager1.registerQuery(makeQuery("q_interval", { type: "interval", intervalMs: 10 }));
  await manager1.ready;

  const completed1 = new Promise((resolve) => {
    manager1.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_interval") resolve(undefined);
    });
  });

  timers.advance(10);
  assert.equal(engine1.calls.length, 1);
  engine1.calls[0].deferred.resolve(makeResult("q_interval"));
  await completed1;
  await Promise.resolve(); // allow lastRunAtMs persistence

  const persisted1 = await stateStore1.load();
  assert.deepEqual(persisted1.q_interval.policy, { type: "interval", intervalMs: 10 });
  assert.equal(persisted1.q_interval.lastRunAtMs, 10);

  timers.advance(5); // now = 15
  manager1.dispose();

  const stateStore2 = createPowerQueryRefreshStateStore({ workbookId: "wb_refresh", storage });
  const engine2 = new ControlledEngine();
  const manager2 = new RefreshManager({
    engine: engine2,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    stateStore: stateStore2,
  });

  const query2 = makeQuery("q_interval");
  delete query2.refreshPolicy;
  manager2.registerQuery(query2);
  await manager2.ready;

  const completed2 = new Promise((resolve) => {
    manager2.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_interval") resolve(undefined);
    });
  });

  timers.advance(4);
  assert.equal(engine2.calls.length, 0);

  timers.advance(1);
  assert.equal(engine2.calls.length, 1, "expected interval schedule to be relative to persisted lastRunAtMs");
  engine2.calls[0].deferred.resolve(makeResult("q_interval"));
  await completed2;

  manager2.dispose();
});

test("RefreshManager persists cron schedules via the desktop refresh state store", async () => {
  const storage = new MapStorage();
  const stateStore1 = createPowerQueryRefreshStateStore({ workbookId: "wb_cron", storage });
  const timers = new FakeTimers(0);

  const engine1 = new ControlledEngine();
  const manager1 = new RefreshManager({
    engine: engine1,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    timezone: "utc",
    stateStore: stateStore1,
  });

  manager1.registerQuery(makeQuery("q_cron", { type: "cron", cron: "* * * * *" }));
  await manager1.ready;

  const completed1 = new Promise((resolve) => {
    manager1.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_cron") resolve(undefined);
    });
  });

  timers.advance(60 * 1000);
  assert.equal(engine1.calls.length, 1);
  engine1.calls[0].deferred.resolve(makeResult("q_cron"));
  await completed1;
  await Promise.resolve();

  const persisted = await stateStore1.load();
  assert.deepEqual(persisted.q_cron.policy, { type: "cron", cron: "* * * * *" });
  assert.equal(persisted.q_cron.lastRunAtMs, 60 * 1000);

  manager1.dispose();

  // Simulate a clock reset (or VM suspend) that would otherwise schedule an
  // already-executed cron minute.
  const resetTimers = new FakeTimers(0);
  const stateStore2 = createPowerQueryRefreshStateStore({ workbookId: "wb_cron", storage });
  const engine2 = new ControlledEngine();
  const manager2 = new RefreshManager({
    engine: engine2,
    concurrency: 1,
    timers: { setTimeout: (...args) => resetTimers.setTimeout(...args), clearTimeout: (id) => resetTimers.clearTimeout(id) },
    now: () => resetTimers.now,
    timezone: "utc",
    stateStore: stateStore2,
  });

  const query2 = makeQuery("q_cron");
  delete query2.refreshPolicy;
  manager2.registerQuery(query2);
  await manager2.ready;

  resetTimers.advance(60 * 1000);
  assert.equal(engine2.calls.length, 0, "expected restored lastRunAtMs to avoid rerunning the same cron minute");

  resetTimers.advance(60 * 1000);
  assert.equal(engine2.calls.length, 1);
  engine2.calls[0].deferred.resolve(makeResult("q_cron"));
  await engine2.calls[0].deferred.promise;

  manager2.dispose();
});

test("RefreshManager saves policy updates to the desktop refresh state store", async () => {
  const storage = new MapStorage();
  const stateStore = createPowerQueryRefreshStateStore({ workbookId: "wb_policy", storage });
  const engine = new ControlledEngine();
  const timers = new FakeTimers();
  const manager = new RefreshManager({
    engine,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    stateStore,
  });

  manager.registerQuery(makeQuery("q_policy", { type: "manual" }));
  await manager.ready;

  manager.registerQuery(makeQuery("q_policy", { type: "interval", intervalMs: 123 }));
  // `RefreshManager` persists state asynchronously and batches overlapping saves.
  // Yield a couple ticks to allow the queued save to flush.
  await Promise.resolve();
  await Promise.resolve();

  const persisted = await stateStore.load();
  assert.deepEqual(persisted.q_policy.policy, { type: "interval", intervalMs: 123 });

  manager.dispose();
});

test("RefreshManager persistence failures are best-effort", async () => {
  const brokenStore = {
    async load() {
      throw new Error("load failed");
    },
    async save() {
      throw new Error("save failed");
    },
  };

  const engine = new ControlledEngine();
  const manager = new RefreshManager({ engine, concurrency: 1, stateStore: brokenStore });

  const query = makeQuery("q_broken", { type: "manual" });
  manager.registerQuery(query);
  await manager.ready;

  const handle = manager.refresh(query.id);
  assert.equal(engine.calls.length, 1);
  engine.calls[0].deferred.resolve(makeResult(query.id));
  await handle.promise;

  manager.dispose();
});
