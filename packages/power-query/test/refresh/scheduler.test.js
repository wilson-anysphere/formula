import assert from "node:assert/strict";
import test from "node:test";

import { RefreshManager } from "../../src/refresh.js";
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
    /** @type {{ queryId: string, deferred: ReturnType<typeof deferred>, signal?: AbortSignal, onProgress?: any }[]} */
    this.calls = [];
  }

  executeQueryWithMeta(query, _context, options) {
    const d = deferred();
    this.calls.push({ queryId: query.id, deferred: d, signal: options?.signal, onProgress: options?.onProgress });
    options?.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey: "k" });
    options?.signal?.addEventListener("abort", () => {
      const err = new Error("Aborted");
      err.name = "AbortError";
      d.reject(err);
    });
    return d.promise;
  }
}

function makeQuery(id, refreshPolicy = { type: "manual" }) {
  return {
    id,
    name: id,
    source: { type: "range", range: { values: [["Value"], [1]], hasHeaders: true } },
    steps: [],
    refreshPolicy,
  };
}

test("RefreshManager: FIFO ordering + concurrency limit", async () => {
  const engine = new ControlledEngine();
  const manager = new RefreshManager({ engine, concurrency: 1 });
  manager.registerQuery(makeQuery("q1"));
  manager.registerQuery(makeQuery("q2"));

  /** @type {string[]} */
  const started = [];
  const unsub = manager.onEvent((evt) => {
    if (evt.type === "started") started.push(evt.job.queryId);
  });

  const h1 = manager.refresh("q1");
  const h2 = manager.refresh("q2");

  assert.deepEqual(started, ["q1"]);
  assert.equal(engine.calls.length, 1);

  engine.calls[0].deferred.resolve(makeResult("q1"));
  await h1.promise;

  assert.deepEqual(started, ["q1", "q2"]);
  assert.equal(engine.calls.length, 2);

  engine.calls[1].deferred.resolve(makeResult("q2"));
  await h2.promise;

  unsub();
  manager.dispose();
});

test("RefreshManager: cancellation of queued job prevents execution", async () => {
  const engine = new ControlledEngine();
  const manager = new RefreshManager({ engine, concurrency: 1 });
  manager.registerQuery(makeQuery("q1"));
  manager.registerQuery(makeQuery("q2"));

  const h1 = manager.refresh("q1");
  const h2 = manager.refresh("q2");

  assert.equal(engine.calls.length, 1);
  h2.cancel();

  await assert.rejects(h2.promise, (err) => err?.name === "AbortError");

  engine.calls[0].deferred.resolve(makeResult("q1"));
  await h1.promise;

  assert.equal(engine.calls.length, 1, "cancelled queued job should never reach the engine");
  manager.dispose();
});

test("RefreshManager: cancellation aborts an in-flight refresh and emits progress", async () => {
  const engine = new ControlledEngine();
  const manager = new RefreshManager({ engine, concurrency: 1 });
  manager.registerQuery(makeQuery("q1"));

  /** @type {string[]} */
  const events = [];
  manager.onEvent((evt) => {
    if (evt.type === "progress") events.push(evt.event.type);
    if (evt.type === "cancelled") events.push("cancelled");
  });

  const handle = manager.refresh("q1");
  assert.equal(engine.calls.length, 1);

  handle.cancel();
  await assert.rejects(handle.promise, (err) => err?.name === "AbortError");
  assert.ok(events.includes("cache:miss"), "engine progress should be forwarded");
  assert.ok(events.includes("cancelled"), "manager should emit cancelled events");
  manager.dispose();
});

class FakeTimers {
  constructor() {
    this.now = 0;
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

test("RefreshManager: interval policy schedules refreshes via injected timers", async () => {
  const engine = new ControlledEngine();
  const timers = new FakeTimers();
  const manager = new RefreshManager({
    engine,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
  });

  manager.registerQuery(makeQuery("q_interval", { type: "interval", intervalMs: 10 }));
  const completed = new Promise((resolve) => {
    manager.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_interval") resolve(undefined);
    });
  });

  assert.equal(engine.calls.length, 0);
  timers.advance(9);
  assert.equal(engine.calls.length, 0);

  timers.advance(1);
  assert.equal(engine.calls.length, 1);
  engine.calls[0].deferred.resolve(makeResult("q_interval"));
  await completed;
  manager.dispose();
});
