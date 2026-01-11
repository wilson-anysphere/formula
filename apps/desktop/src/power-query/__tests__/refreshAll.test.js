import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../../../../packages/power-query/src/table.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryRefreshOrchestrator } from "../refreshAll.ts";

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

function abortError() {
  const err = new Error("Aborted");
  err.name = "AbortError";
  return err;
}

async function waitForAbort(signal) {
  if (signal?.aborted) throw abortError();
  await new Promise((_, reject) => {
    signal?.addEventListener(
      "abort",
      () => {
        reject(abortError());
      },
      { once: true },
    );
  });
}

class ScriptedEngine {
  /**
   * @param {Record<string, { table?: DataTable, error?: Error, waitForAbort?: boolean }>} scripts
   */
  constructor(scripts) {
    this.scripts = scripts;
    /** @type {string[]} */
    this.calls = [];
    /** @type {string[]} */
    this.aborted = [];
  }

  createSession(options = {}) {
    return { credentialCache: new Map(), permissionCache: new Map(), now: options.now };
  }

  async executeQueryWithMetaInSession(query, _context, options) {
    this.calls.push(query.id);

    const script = this.scripts[query.id] ?? {};

    if (options?.signal?.aborted) throw abortError();
    if (script.waitForAbort) {
      try {
        await waitForAbort(options?.signal);
      } finally {
        this.aborted.push(query.id);
      }
    }

    if (script.error) throw script.error;

    const table = script.table ?? DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
    return { table, meta: makeMeta(query.id, table) };
  }
}

test("DesktopPowerQueryRefreshOrchestrator refreshes dependencies once and applies completed targets", async () => {
  const tableRef = DataTable.fromGrid([["RefCol"], ["r1"]], { hasHeaders: true, inferTypes: true });
  const tableOps = DataTable.fromGrid([["OpCol"], ["o1"]], { hasHeaders: true, inferTypes: true });
  const tableShared = DataTable.fromGrid([["SharedCol"], ["s1"]], { hasHeaders: true, inferTypes: true });

  const engine = new ScriptedEngine({
    q_shared: { table: tableShared },
    q_ref: { table: tableRef },
    q_ops: { table: tableOps },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 1, batchSize: 1 });

  const qShared = {
    id: "q_shared",
    name: "Shared",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };

  const qRef = {
    id: "q_ref",
    name: "Ref",
    source: { type: "query", queryId: "q_shared" },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  const qOps = {
    id: "q_ops",
    name: "Ops",
    source: { type: "range", range: { values: [["Y"], [2]], hasHeaders: true } },
    steps: [
      {
        id: "s1",
        name: "merge",
        operation: { type: "merge", rightQuery: "q_shared", joinType: "inner", leftKey: "X", rightKey: "X" },
      },
      { id: "s2", name: "append", operation: { type: "append", queries: ["q_shared", "q_ref"] } },
    ],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 3 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  orchestrator.registerQuery(qShared);
  orchestrator.registerQuery(qRef);
  orchestrator.registerQuery(qOps);

  const applied = new Promise((resolve) => {
    /** @type {Set<string>} */
    const done = new Set();
    const unsub = orchestrator.onEvent((evt) => {
      if (evt.type === "apply:completed" && (evt.queryId === "q_ref" || evt.queryId === "q_ops")) {
        done.add(evt.queryId);
        if (done.size === 2) {
          unsub();
          resolve(done);
        }
      }
    });
  });

  const handle = orchestrator.refreshAll(["q_ref", "q_ops"]);
  await handle.promise;
  await applied;

  assert.deepEqual(engine.calls, ["q_shared", "q_ref", "q_ops"]);
  assert.equal(engine.calls.filter((id) => id === "q_shared").length, 1);

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "RefCol");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, "r1");

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 3 }).value, "OpCol");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 3 }).value, "o1");

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator serializes apply operations to avoid nested document batches", async () => {
  const table1 = DataTable.fromGrid(
    [["A"], ...Array.from({ length: 10 }, (_, idx) => [idx + 1])],
    { hasHeaders: true, inferTypes: true },
  );
  const table2 = DataTable.fromGrid(
    [["B"], ...Array.from({ length: 3 }, (_, idx) => [`v${idx + 1}`])],
    { hasHeaders: true, inferTypes: true },
  );

  const engine = new ScriptedEngine({
    q1: { table: table1 },
    q2: { table: table2 },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  // Concurrency > 1 means refresh completion events can interleave, so apply must be queued.
  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 2, batchSize: 1 });

  const q1 = {
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  const q2 = {
    id: "q2",
    name: "Q2",
    source: { type: "range", range: { values: [["Y"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 3 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  orchestrator.registerQuery(q1);
  orchestrator.registerQuery(q2);

  const applied = new Promise((resolve) => {
    /** @type {Set<string>} */
    const done = new Set();
    const unsub = orchestrator.onEvent((evt) => {
      if (evt.type === "apply:completed" && (evt.queryId === "q1" || evt.queryId === "q2")) {
        done.add(evt.queryId);
        if (done.size === 2) {
          unsub();
          resolve(done);
        }
      }
    });
  });

  const handle = orchestrator.refreshAll(["q1", "q2"]);
  await handle.promise;
  await applied;

  // Each apply should be its own batch/undo entry. Without apply serialization, concurrent
  // applies would nest and collapse into a single history entry.
  assert.equal(doc.history.length, 2);
  assert.equal(doc.batchDepth, 0);

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "A");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, 1);

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 3 }).value, "B");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 3 }).value, "v1");

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator shares apply serialization across instances for the same document", async () => {
  const table1 = DataTable.fromGrid(
    [["A"], ...Array.from({ length: 20 }, (_, idx) => [idx + 1])],
    { hasHeaders: true, inferTypes: true },
  );
  const table2 = DataTable.fromGrid([["B"], ["ok"]], { hasHeaders: true, inferTypes: true });

  const engine = new ScriptedEngine({
    q1: { table: table1 },
    q2: { table: table2 },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orch1 = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 1, batchSize: 1 });
  const orch2 = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 1, batchSize: 1 });

  orch1.registerQuery({
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  orch2.registerQuery({
    id: "q2",
    name: "Q2",
    source: { type: "range", range: { values: [["Y"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 3 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  const applied = new Promise((resolve, reject) => {
    /** @type {Set<string>} */
    const done = new Set();
    const maybeDone = () => {
      if (done.size !== 2) return;
      unsub1();
      unsub2();
      resolve(undefined);
    };

    const unsub1 = orch1.onEvent((evt) => {
      if (evt.type === "apply:error") reject(evt.error);
      if (evt.type === "apply:completed" && evt.queryId === "q1") {
        done.add("q1");
        maybeDone();
      }
    });

    const unsub2 = orch2.onEvent((evt) => {
      if (evt.type === "apply:error") reject(evt.error);
      if (evt.type === "apply:completed" && evt.queryId === "q2") {
        done.add("q2");
        maybeDone();
      }
    });
  });

  const handle1 = orch1.refreshAll(["q1"]);

  const started = new Promise((resolve) => {
    const unsub = orch1.onEvent((evt) => {
      if (evt.type === "apply:started" && evt.queryId === "q1") {
        unsub();
        resolve(undefined);
      }
    });
  });

  await started;
  const handle2 = orch2.refreshAll(["q2"]);

  await Promise.all([handle1.promise, handle2.promise]);
  await applied;

  // Each apply should create its own history entry, even when driven by different orchestrator instances.
  assert.equal(doc.history.length, 2);
  assert.equal(doc.batchDepth, 0);

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "A");
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 3 }).value, "B");

  orch1.dispose();
  orch2.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator cancelQuery aborts the apply phase for that query only", async () => {
  const bigTable = DataTable.fromGrid(
    [["A"], ...Array.from({ length: 50 }, (_, i) => [i + 1])],
    { hasHeaders: true, inferTypes: true },
  );
  const smallTable = DataTable.fromGrid([["B"], [1]], { hasHeaders: true, inferTypes: true });

  const engine = new ScriptedEngine({
    q1: { table: bigTable },
    q2: { table: smallTable },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 2, batchSize: 1 });

  orchestrator.registerQuery({
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  orchestrator.registerQuery({
    id: "q2",
    name: "Q2",
    source: { type: "range", range: { values: [["X"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet2", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  const handle = orchestrator.refreshAll(["q1", "q2"]);

  let cancelled = false;
  let q1Cancelled = false;
  let q2Completed = false;

  const done = new Promise((resolve, reject) => {
    const unsub = orchestrator.onEvent((evt) => {
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
  assert.equal(doc.batchDepth, 0);

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator cancelQuery during refresh completion prevents apply from starting", async () => {
  const bigTable = DataTable.fromGrid(
    [["A"], ...Array.from({ length: 20 }, (_, i) => [i + 1])],
    { hasHeaders: true, inferTypes: true },
  );
  const smallTable = DataTable.fromGrid([["B"], [1]], { hasHeaders: true, inferTypes: true });

  const engine = new ScriptedEngine({
    q1: { table: bigTable },
    q2: { table: smallTable },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 2, batchSize: 1 });

  orchestrator.registerQuery({
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  orchestrator.registerQuery({
    id: "q2",
    name: "Q2",
    source: { type: "range", range: { values: [["X"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet2", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  const handle = orchestrator.refreshAll(["q1", "q2"]);

  let cancelled = false;
  let q1ApplyCancelled = false;
  let q2Applied = false;

  const done = new Promise((resolve, reject) => {
    const unsub = orchestrator.onEvent((evt) => {
      if (evt.type === "apply:error") {
        unsub();
        reject(evt.error);
        return;
      }

      if (evt.type === "completed" && evt?.job?.queryId === "q1" && !cancelled) {
        cancelled = true;
        handle.cancelQuery?.("q1");
      }

      if (evt.type === "apply:cancelled" && evt.queryId === "q1") {
        q1ApplyCancelled = true;
      }

      if (evt.type === "apply:completed" && evt.queryId === "q2") {
        q2Applied = true;
      }

      if (q1ApplyCancelled && q2Applied) {
        unsub();
        resolve(undefined);
      }
    });
  });

  await handle.promise;
  await done;

  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.equal(doc.getCell("Sheet2", { row: 0, col: 0 }).value, "B");
  assert.equal(doc.batchDepth, 0);

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator notifies onSuccessfulRun for completed queries", async () => {
  const table = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
  const engine = new ScriptedEngine({ q1: { table } });
  const doc = new DocumentController({ engine: new MockEngine() });

  /** @type {Array<{ queryId: string, completedAtMs: number }>} */
  const calls = [];

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({
    engine,
    document: doc,
    concurrency: 1,
    batchSize: 1,
    onSuccessfulRun: (queryId, completedAtMs) => {
      calls.push({ queryId, completedAtMs });
    },
  });

  orchestrator.registerQuery({
    id: "q1",
    name: "Q1",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  });

  await orchestrator.refreshAll(["q1"]).promise;
  assert.equal(calls.length, 1);
  assert.equal(calls[0].queryId, "q1");
  assert.ok(Number.isFinite(calls[0].completedAtMs));

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator cancels downstream targets on dependency error but continues independent branches", async () => {
  const tableOther = DataTable.fromGrid([["Other"], ["ok"]], { hasHeaders: true, inferTypes: true });

  const engine = new ScriptedEngine({
    q_fail: { error: new Error("boom") },
    q_other: { table: tableOther },
    q_down: { table: DataTable.fromGrid([["Down"], ["no"]], { hasHeaders: true, inferTypes: true }) },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 2, batchSize: 1 });

  const qFail = {
    id: "q_fail",
    name: "Fail",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };
  const qDown = {
    id: "q_down",
    name: "Downstream",
    source: { type: "query", queryId: "q_fail" },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };
  const qOther = {
    id: "q_other",
    name: "Other",
    source: { type: "range", range: { values: [["Y"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 3 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  orchestrator.registerQuery(qFail);
  orchestrator.registerQuery(qDown);
  orchestrator.registerQuery(qOther);

  const outcomes = new Promise((resolve) => {
    /** @type {Set<string>} */
    const seen = new Set();
    const unsub = orchestrator.onEvent((evt) => {
      if (evt.type === "apply:completed" && evt.queryId === "q_other") {
        seen.add("other-applied");
      }
      if (evt.type === "apply:cancelled" && evt.queryId === "q_down") {
        seen.add("down-cancelled");
      }
      if (seen.size === 2) {
        unsub();
        resolve(seen);
      }
    });
  });

  const handle = orchestrator.refreshAll(["q_down", "q_other"]);
  await assert.rejects(handle.promise, (err) => err?.message?.includes("boom"));
  await outcomes;

  assert.ok(engine.calls.includes("q_fail"));
  assert.ok(engine.calls.includes("q_other"));
  assert.ok(!engine.calls.includes("q_down"));

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 3 }).value, "Other");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 3 }).value, "ok");

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, null);

  orchestrator.dispose();
});

test("DesktopPowerQueryRefreshOrchestrator cancel() aborts execution and apply", async () => {
  const largeTable = DataTable.fromGrid(
    [["A"], ...Array.from({ length: 50 }, (_, idx) => [idx + 1])],
    { hasHeaders: true, inferTypes: true },
  );

  const engine = new ScriptedEngine({
    q_long: { waitForAbort: true, table: DataTable.fromGrid([["Long"], ["nope"]], { hasHeaders: true, inferTypes: true }) },
    q_apply: { table: largeTable },
  });
  const doc = new DocumentController({ engine: new MockEngine() });

  const orchestrator = new DesktopPowerQueryRefreshOrchestrator({ engine, document: doc, concurrency: 2, batchSize: 1 });

  const qLong = {
    id: "q_long",
    name: "Long",
    source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 3 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };
  const qApply = {
    id: "q_apply",
    name: "Apply",
    source: { type: "range", range: { values: [["Y"], [2]], hasHeaders: true } },
    steps: [],
    destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, clearExisting: true },
    refreshPolicy: { type: "manual" },
  };

  orchestrator.registerQuery(qLong);
  orchestrator.registerQuery(qApply);

  const handle = orchestrator.refreshAll(["q_apply", "q_long"]);

  const done = new Promise((resolve) => {
    /** @type {Set<string>} */
    const cancelled = new Set();
    let requested = false;
    const unsub = orchestrator.onEvent((evt) => {
      if (evt.type === "apply:progress" && evt.queryId === "q_apply" && !requested) {
        requested = true;
        handle.cancel();
      }
      if (evt.type === "apply:cancelled" && (evt.queryId === "q_apply" || evt.queryId === "q_long")) {
        cancelled.add(evt.queryId);
        if (cancelled.size === 2) {
          unsub();
          resolve(cancelled);
        }
      }
    });
  });

  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  await done;

  // Apply should have been cancelled, leaving the sheet untouched.
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, null);
  assert.equal(doc.getUsedRange("Sheet1"), null);
  assert.equal(doc.batchDepth, 0);

  assert.ok(engine.calls.includes("q_long"));
  assert.ok(engine.aborted.includes("q_long"));

  orchestrator.dispose();
});
