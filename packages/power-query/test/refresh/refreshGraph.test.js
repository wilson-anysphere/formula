import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { RefreshOrchestrator } from "../../src/refreshGraph.js";
import { DataTable } from "../../src/table.js";

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

function makeResult(queryId) {
  const table = new DataTable([], []);
  return {
    table,
    meta: {
      queryId,
      startedAt: new Date(0),
      completedAt: new Date(0),
      refreshedAt: new Date(0),
      sources: [],
      outputSchema: { columns: [], inferred: true },
      outputRowCount: 0,
    },
  };
}

class ControlledEngine {
  constructor() {
    /** @type {{ queryId: string, deferred: ReturnType<typeof deferred>, signal?: AbortSignal }[]} */
    this.calls = [];
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

function makeQuery(id, source, steps = []) {
  return {
    id,
    name: id,
    source,
    steps,
  };
}

test("RefreshOrchestrator: DAG ordering (B depends on A)", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  const handle = orchestrator.refreshAll(["B"]);

  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A", "dependency should be refreshed first");

  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2);
  assert.equal(engine.calls[1].queryId, "B");

  engine.calls[1].deferred.resolve(makeResult("B"));
  await handle.promise;
});

test("RefreshOrchestrator: merge dependency ordering (B merge depends on A)", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }, [
      {
        id: "merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "A", joinType: "left", leftKey: "Key", rightKey: "Key" },
      },
    ]),
  );

  const handle = orchestrator.refreshAll(["B"]);

  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A");

  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2);
  assert.equal(engine.calls[1].queryId, "B");

  engine.calls[1].deferred.resolve(makeResult("B"));
  await handle.promise;
});

test("RefreshOrchestrator: append dependency ordering (B append depends on A)", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }, [
      {
        id: "append",
        name: "Append",
        operation: { type: "append", queries: ["A"] },
      },
    ]),
  );

  const handle = orchestrator.refreshAll(["B"]);

  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A");

  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2);
  assert.equal(engine.calls[1].queryId, "B");

  engine.calls[1].deferred.resolve(makeResult("B"));
  await handle.promise;
});

test("RefreshOrchestrator: dedup shared dependency results (A only loads once)", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async (_path) => {
        reads += 1;
        return "Value\n1\n";
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "csv", path: "file.csv" }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  await orchestrator.refreshAll(["A", "B"]).promise;
  assert.equal(reads, 1, "dependency query should not be re-executed inside dependents");
});

test("RefreshOrchestrator: shared execution session dedupes credential prompts", async () => {
  let credentialRequests = 0;
  const engine = new QueryEngine({
    fileAdapter: { readText: async (_path) => "Value\n1\n" },
    onCredentialRequest: async () => {
      credentialRequests += 1;
      return { token: "ok" };
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("Q1", { type: "csv", path: "file.csv" }));
  orchestrator.registerQuery(makeQuery("Q2", { type: "csv", path: "file.csv" }));

  await orchestrator.refreshAll(["Q1", "Q2"]).promise;
  assert.equal(credentialRequests, 1);
});

test("RefreshOrchestrator: cycle detection emits a clear error", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "query", queryId: "B" }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["A"]);
  await assert.rejects(handle.promise, /cycle/i);
  assert.equal(engine.calls.length, 0, "cycle should be detected before any engine work starts");

  const errEvt = events.find((e) => e.type === "error");
  assert.ok(errEvt, "orchestrator should emit an error event");
  assert.match(String(errEvt.error?.message ?? ""), /A -> B -> A/);
});

test("RefreshOrchestrator: unknown target query emits error event and rejects", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["Missing"]);
  await assert.rejects(handle.promise, /Unknown query 'Missing'/);
  assert.equal(engine.calls.length, 0);

  const errEvt = events.find((e) => e.type === "error");
  assert.ok(errEvt, "orchestrator should emit an error event");
  assert.equal(errEvt.job.queryId, "Missing");
});

test("RefreshOrchestrator: missing dependency emits error event and rejects", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "Missing" }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["B"]);
  await assert.rejects(handle.promise, /Unknown query 'Missing' \(dependency of 'B'\)/);
  assert.equal(engine.calls.length, 0);

  const errEvt = events.find((e) => e.type === "error");
  assert.ok(errEvt, "orchestrator should emit an error event");
  assert.equal(errEvt.job.queryId, "B", "error should be associated with the query that references the missing dependency");
});

test("RefreshOrchestrator: cancellation stops remaining queued jobs", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("q2", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }));

  /** @type {string[]} */
  const cancelled = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "cancelled") cancelled.push(evt.job.queryId);
  });

  const handle = orchestrator.refreshAll(["q1", "q2"]);

  assert.equal(engine.calls.length, 1);
  assert.ok(["q1", "q2"].includes(engine.calls[0].queryId));

  handle.cancel();
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  assert.equal(engine.calls.length, 1, "cancelled session should not start queued jobs");
  assert.deepEqual(new Set(cancelled), new Set(["q1", "q2"]), "cancellation should emit per-job cancelled events");
});

test("RefreshOrchestrator: cancellation rejects even when dependents were never scheduled", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("C", { type: "query", queryId: "B" }));

  /** @type {string[]} */
  const cancelled = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "cancelled") cancelled.push(evt.job.queryId);
  });

  const handle = orchestrator.refreshAll(["C"]);
  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A");

  handle.cancel();
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  assert.deepEqual(new Set(cancelled), new Set(["A", "B", "C"]));
});
