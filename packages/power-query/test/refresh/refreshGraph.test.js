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
    options?.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey: "k" });
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

test("RefreshOrchestrator: supports query ids like '__proto__' without prototype pollution", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("__proto__", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  const handle = orchestrator.refreshAll(["__proto__"]);
  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "__proto__");

  engine.calls[0].deferred.resolve(makeResult("__proto__"));
  const results = await handle.promise;

  assert.equal(Object.getPrototypeOf(results), Object.prototype);
  assert.equal(Object.prototype.hasOwnProperty.call(results, "__proto__"), true);
  assert.equal(results["__proto__"].meta.queryId, "__proto__");
  // Ensure we did not mutate the global object prototype.
  assert.equal(({}).polluted, undefined);
});

test("RefreshOrchestrator: query references work when the dependency id is '__proto__'", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return "Value\n1\n";
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("__proto__", { type: "csv", path: "file.csv", options: { hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "__proto__" }));

  const results = await orchestrator.refreshAll(["B"]).promise;
  assert.equal(reads, 1);
  assert.deepEqual(Object.keys(results), ["B"]);
});

test("RefreshOrchestrator: merge step reuses dependency when rightQuery id is '__proto__'", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return ["Key,Value", "1,a", "2,b"].join("\n");
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("__proto__", { type: "csv", path: "file.csv", options: { hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }, [
      { id: "merge", name: "Merge", operation: { type: "merge", rightQuery: "__proto__", joinType: "left", leftKey: "Key", rightKey: "Key" } },
    ]),
  );

  await orchestrator.refreshAll(["B"]).promise;
  assert.equal(reads, 1);
});

test("RefreshOrchestrator: append step reuses dependencies when appended query id is '__proto__'", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return "Value\n1\n";
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("__proto__", { type: "csv", path: "file.csv", options: { hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }, [
      { id: "append", name: "Append", operation: { type: "append", queries: ["__proto__"] } },
    ]),
  );

  await orchestrator.refreshAll(["B"]).promise;
  assert.equal(reads, 1);
});

test("RefreshOrchestrator: refresh(queryId) refreshes dependencies and returns the target result", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  const handle = orchestrator.refresh("B");
  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A");
  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2);
  assert.equal(engine.calls[1].queryId, "B");
  engine.calls[1].deferred.resolve(makeResult("B"));

  const result = await handle.promise;
  assert.equal(result.meta.queryId, "B");
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

test("RefreshOrchestrator: merge dependency does not re-execute rightQuery inside merge step", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return ["Key,Value", "1,a", "2,b"].join("\n");
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "csv", path: "file.csv", options: { hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }, [
      { id: "merge", name: "Merge", operation: { type: "merge", rightQuery: "A", joinType: "left", leftKey: "Key", rightKey: "Key" } },
    ]),
  );

  await orchestrator.refreshAll(["B"]).promise;
  assert.equal(reads, 1, "merge step should reuse precomputed dependency result");
});

test("RefreshOrchestrator: events include sessionId + dependency/target phase", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  /** @type {{ queryId: string; sessionId: string; phase: string }[]} */
  const started = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "started") {
      started.push({ queryId: evt.job.queryId, sessionId: evt.sessionId, phase: evt.phase });
    }
  });

  const handle = orchestrator.refreshAll(["B"]);

  assert.equal(engine.calls.length, 1);
  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2);
  engine.calls[1].deferred.resolve(makeResult("B"));
  await handle.promise;

  const byQuery = Object.fromEntries(started.map((e) => [e.queryId, e]));
  assert.equal(byQuery["A"]?.phase, "dependency");
  assert.equal(byQuery["B"]?.phase, "target");
  assert.equal(byQuery["A"]?.sessionId, handle.sessionId);
  assert.equal(byQuery["B"]?.sessionId, handle.sessionId);
});

test("RefreshOrchestrator: forwards progress events with sessionId and phase", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  /** @type {{ queryId: string; sessionId: string; phase: string; eventType: string }[]} */
  const progress = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "progress") {
      progress.push({ queryId: evt.job.queryId, sessionId: evt.sessionId, phase: evt.phase, eventType: evt.event.type });
    }
  });

  const handle = orchestrator.refreshAll(["B"]);
  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));
  engine.calls[1].deferred.resolve(makeResult("B"));
  await handle.promise;

  const byQuery = Object.fromEntries(progress.map((e) => [e.queryId, e]));
  assert.equal(byQuery["A"]?.eventType, "cache:miss");
  assert.equal(byQuery["A"]?.phase, "dependency");
  assert.equal(byQuery["B"]?.eventType, "cache:miss");
  assert.equal(byQuery["B"]?.phase, "target");
  assert.equal(byQuery["A"]?.sessionId, handle.sessionId);
  assert.equal(byQuery["B"]?.sessionId, handle.sessionId);
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

test("RefreshOrchestrator: append dependencies do not re-execute appended queries inside append step", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return ["Value", "1"].join("\n");
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "csv", path: "file.csv", options: { hasHeaders: true } }));
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }, [
      { id: "append", name: "Append", operation: { type: "append", queries: ["A"] } },
    ]),
  );

  await orchestrator.refreshAll(["B"]).promise;
  assert.equal(reads, 1, "append step should reuse precomputed dependency result");
});

test("RefreshOrchestrator: runs independent queries concurrently", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("q2", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }));

  const handle = orchestrator.refreshAll(["q1", "q2"]);

  assert.equal(engine.calls.length, 2, "both root queries should start immediately when concurrency allows");
  const started = new Set(engine.calls.map((c) => c.queryId));
  assert.deepEqual(started, new Set(["q1", "q2"]));

  for (const call of engine.calls) call.deferred.resolve(makeResult(call.queryId));
  await handle.promise;
});

test("RefreshOrchestrator: refreshAll sessions get unique sessionIds", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  const h1 = orchestrator.refreshAll(["q1"]);
  engine.calls[0].deferred.resolve(makeResult("q1"));
  await h1.promise;

  const h2 = orchestrator.refreshAll(["q1"]);
  engine.calls[1].deferred.resolve(makeResult("q1"));
  await h2.promise;

  assert.notEqual(h1.sessionId, h2.sessionId);
});

test("RefreshOrchestrator: triggerOnOpen refreshes on-open queries with dependencies", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery({
    ...makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }),
    refreshPolicy: { type: "manual" },
  });
  orchestrator.registerQuery({
    ...makeQuery("B", { type: "query", queryId: "A" }),
    refreshPolicy: { type: "on-open" },
  });
  orchestrator.registerQuery({
    ...makeQuery("C", { type: "range", range: { values: [["Value"], [3]], hasHeaders: true } }),
    refreshPolicy: { type: "on-open" },
  });

  const handle = orchestrator.triggerOnOpen();
  assert.equal(engine.calls.length, 2);
  assert.deepEqual(new Set(engine.calls.map((c) => c.queryId)), new Set(["A", "C"]));

  const callA = engine.calls.find((c) => c.queryId === "A");
  const callC = engine.calls.find((c) => c.queryId === "C");
  assert.ok(callA);
  assert.ok(callC);
  callA.deferred.resolve(makeResult("A"));
  callC.deferred.resolve(makeResult("C"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 3);
  assert.equal(engine.calls[2].queryId, "B");
  engine.calls[2].deferred.resolve(makeResult("B"));

  const results = await handle.promise;
  assert.deepEqual(new Set(Object.keys(results)), new Set(["B", "C"]));
});

test("RefreshOrchestrator: triggerOnOpen(queryId) is a no-op when the query is not on-open", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery({ ...makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }), refreshPolicy: { type: "manual" } });

  const handle = orchestrator.triggerOnOpen("A");
  const results = await handle.promise;
  assert.equal(engine.calls.length, 0);
  assert.deepEqual(results, {});
});

test("RefreshOrchestrator: job ids are namespaced by the refreshAll sessionId", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["q1"]);
  engine.calls[0].deferred.resolve(makeResult("q1"));
  await handle.promise;

  const started = events.find((e) => e.type === "started" && e.job.queryId === "q1");
  assert.ok(started);
  assert.equal(typeof started.job.id, "string");
  assert.ok(started.job.id.startsWith(`${handle.sessionId}:`));
});

test("RefreshOrchestrator: refreshAll([]) resolves immediately without engine work", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  const handle = orchestrator.refreshAll([]);
  const results = await handle.promise;
  assert.deepEqual(results, {});
  assert.equal(engine.calls.length, 0);
});

test("RefreshOrchestrator: refreshAll propagates reason to jobs", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
  orchestrator.registerQuery(makeQuery("q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  /** @type {any[]} */
  const started = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "started") started.push(evt);
  });

  const handle = orchestrator.refreshAll(["q1"], "cron");
  engine.calls[0].deferred.resolve(makeResult("q1"));
  await handle.promise;

  assert.equal(started.length, 1);
  assert.equal(started[0].job.reason, "cron");
});

test("RefreshOrchestrator: shared dependency executes once when not explicitly targeted", async () => {
  let reads = 0;
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        reads += 1;
        return "Value\n1\n";
      },
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "csv", path: "file.csv" }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("C", { type: "query", queryId: "A" }));

  const results = await orchestrator.refreshAll(["B", "C"]).promise;
  assert.equal(reads, 1);
  assert.deepEqual(new Set(Object.keys(results)), new Set(["B", "C"]));
});

test("RefreshOrchestrator: fan-out dependencies schedule dependents concurrently", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("C", { type: "query", queryId: "A" }));

  const handle = orchestrator.refreshAll(["B", "C"]);

  assert.equal(engine.calls.length, 1);
  assert.equal(engine.calls[0].queryId, "A");

  engine.calls[0].deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 3);
  assert.deepEqual(
    new Set(engine.calls.slice(1).map((c) => c.queryId)),
    new Set(["B", "C"]),
    "both dependents should start once the shared dependency completes",
  );

  engine.calls[1].deferred.resolve(makeResult(engine.calls[1].queryId));
  engine.calls[2].deferred.resolve(makeResult(engine.calls[2].queryId));

  const results = await handle.promise;
  assert.deepEqual(new Set(Object.keys(results)), new Set(["B", "C"]));
});

test("RefreshOrchestrator: refreshAll() with no queryIds refreshes all registered queries", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("C", { type: "range", range: { values: [["Value"], [3]], hasHeaders: true } }));

  const handle = orchestrator.refreshAll();

  assert.equal(engine.calls.length, 2);
  assert.deepEqual(new Set(engine.calls.map((c) => c.queryId)), new Set(["A", "C"]));

  const callA = engine.calls.find((c) => c.queryId === "A");
  assert.ok(callA);
  callA.deferred.resolve(makeResult("A"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 3);
  const callB = engine.calls.find((c) => c.queryId === "B");
  assert.ok(callB, "dependent query should be scheduled after its dependency completes");

  const callC = engine.calls.find((c) => c.queryId === "C");
  assert.ok(callC);
  callB.deferred.resolve(makeResult("B"));
  callC.deferred.resolve(makeResult("C"));
  const results = await handle.promise;
  assert.deepEqual(new Set(Object.keys(results)), new Set(["A", "B", "C"]));
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

test("RefreshOrchestrator: shared execution session dedupes permission prompts", async () => {
  let permissionRequests = 0;
  const engine = new QueryEngine({
    fileAdapter: { readText: async (_path) => "Value\n1\n" },
    onPermissionRequest: async () => {
      permissionRequests += 1;
      return true;
    },
  });

  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("Q1", { type: "csv", path: "file.csv" }));
  orchestrator.registerQuery(makeQuery("Q2", { type: "csv", path: "file.csv" }));

  await orchestrator.refreshAll(["Q1", "Q2"]).promise;
  assert.equal(permissionRequests, 1);
});

test("RefreshOrchestrator: supports injecting a shared now() for the session", async () => {
  const engine = new QueryEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1, now: () => 123 });
  orchestrator.registerQuery(makeQuery("Q1", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  const results = await orchestrator.refreshAll(["Q1"]).promise;
  assert.equal(results["Q1"].meta.startedAt.getTime(), 123);
  assert.equal(results["Q1"].meta.completedAt.getTime(), 123);
  assert.equal(results["Q1"].meta.refreshedAt.getTime(), 123);
});

test("RefreshOrchestrator: injected now() is used for graph error timestamps", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1, now: () => 123 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["Missing"]);
  await assert.rejects(handle.promise, /Unknown query 'Missing'/);

  const errEvt = events.find((e) => e.type === "error");
  assert.ok(errEvt);
  assert.equal(errEvt.job.queuedAt.getTime(), 123);
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

test("RefreshOrchestrator: cycle detection includes merge dependencies", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(
    makeQuery("A", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }, [
      { id: "merge", name: "Merge", operation: { type: "merge", rightQuery: "B", joinType: "left", leftKey: "Key", rightKey: "Key" } },
    ]),
  );
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Key"], [1]], hasHeaders: true } }, [
      { id: "merge", name: "Merge", operation: { type: "merge", rightQuery: "A", joinType: "left", leftKey: "Key", rightKey: "Key" } },
    ]),
  );

  const handle = orchestrator.refreshAll(["A"]);
  await assert.rejects(handle.promise, /cycle/i);
  assert.equal(engine.calls.length, 0);
});

test("RefreshOrchestrator: cycle detection includes append dependencies", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(
    makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }, [
      { id: "append", name: "Append", operation: { type: "append", queries: ["B"] } },
    ]),
  );
  orchestrator.registerQuery(
    makeQuery("B", { type: "range", range: { values: [["Value"], [2]], hasHeaders: true } }, [
      { id: "append", name: "Append", operation: { type: "append", queries: ["A"] } },
    ]),
  );

  const handle = orchestrator.refreshAll(["A"]);
  await assert.rejects(handle.promise, /cycle/i);
  assert.equal(engine.calls.length, 0);
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

test("RefreshOrchestrator: cancelQuery cancels a branch without stopping independent targets", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("D", { type: "range", range: { values: [["Value"], [4]], hasHeaders: true } }));

  /** @type {string[]} */
  const completed = [];
  /** @type {string[]} */
  const cancelled = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "completed") completed.push(evt.job.queryId);
    if (evt.type === "cancelled") cancelled.push(evt.job.queryId);
  });

  const handle = orchestrator.refreshAll(["B", "D"]);
  assert.equal(engine.calls.length, 2);

  const callA = engine.calls.find((c) => c.queryId === "A");
  const callD = engine.calls.find((c) => c.queryId === "D");
  assert.ok(callA);
  assert.ok(callD);

  handle.cancelQuery?.("A");
  await new Promise((r) => setImmediate(r));

  assert.ok(cancelled.includes("B"), "dependent target should be cancelled after cancelling its dependency");
  assert.equal(engine.calls.length, 2, "cancelling a dependency should not start dependents");

  callD.deferred.resolve(makeResult("D"));
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  assert.ok(completed.includes("D"), "independent target should still complete");
});

test("RefreshOrchestrator: injected now() is used for synthetic cancelled timestamps", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1, now: () => 123 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));

  /** @type {any[]} */
  const events = [];
  orchestrator.onEvent((evt) => events.push(evt));

  const handle = orchestrator.refreshAll(["B"]);
  assert.equal(engine.calls.length, 1);
  handle.cancel();
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");

  const cancelledB = events.find((e) => e.type === "cancelled" && e.job.queryId === "B");
  assert.ok(cancelledB);
  assert.equal(cancelledB.job.queuedAt.getTime(), 123);
  assert.equal(cancelledB.job.completedAt.getTime(), 123);
});

test("RefreshOrchestrator: dependency error rejects and cancels unscheduled dependents", async () => {
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

  engine.calls[0].deferred.reject(new Error("Boom"));
  await assert.rejects(handle.promise, /Boom/);
  assert.deepEqual(new Set(cancelled), new Set(["B", "C"]));
});

test("RefreshOrchestrator: dependency errors do not cancel independent targets", async () => {
  const engine = new ControlledEngine();
  const orchestrator = new RefreshOrchestrator({ engine, concurrency: 2 });
  orchestrator.registerQuery(makeQuery("A", { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } }));
  orchestrator.registerQuery(makeQuery("B", { type: "query", queryId: "A" }));
  orchestrator.registerQuery(makeQuery("D", { type: "range", range: { values: [["Value"], [4]], hasHeaders: true } }));

  /** @type {string[]} */
  const completed = [];
  /** @type {string[]} */
  const cancelled = [];
  orchestrator.onEvent((evt) => {
    if (evt.type === "completed") completed.push(evt.job.queryId);
    if (evt.type === "cancelled") cancelled.push(evt.job.queryId);
  });

  const handle = orchestrator.refreshAll(["B", "D"]);
  assert.equal(engine.calls.length, 2);
  assert.deepEqual(new Set(engine.calls.map((c) => c.queryId)), new Set(["A", "D"]));

  const callA = engine.calls.find((c) => c.queryId === "A");
  const callD = engine.calls.find((c) => c.queryId === "D");
  assert.ok(callA);
  assert.ok(callD);

  callA.deferred.reject(new Error("Boom"));
  await new Promise((r) => setImmediate(r));

  assert.equal(engine.calls.length, 2, "dependent query should not be started after dependency error");
  assert.ok(cancelled.includes("B"), "dependent target should be cancelled");

  callD.deferred.resolve(makeResult("D"));
  await assert.rejects(handle.promise, /Boom/);
  assert.ok(completed.includes("D"), "independent target should still complete");
});
