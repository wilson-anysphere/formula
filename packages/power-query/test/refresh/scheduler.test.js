import assert from "node:assert/strict";
import test from "node:test";

import { nextCronRun, parseCronExpression } from "../../src/cron.js";
import { RefreshManager } from "../../src/refresh.js";
import { InMemoryRefreshStateStore } from "../../src/refreshStateStore.js";
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

test("RefreshManager: supports query ids like '__proto__' in persisted state", async () => {
  const store = new InMemoryRefreshStateStore();
  const engine = new ControlledEngine();
  const manager = new RefreshManager({ engine, concurrency: 1, stateStore: store, now: () => 0 });

  manager.registerQuery(makeQuery("__proto__", { type: "interval", intervalMs: 10 }));
  await manager.ready;

  const handle = manager.refresh("__proto__");
  assert.equal(engine.calls.length, 1);
  engine.calls[0].deferred.resolve(makeResult("__proto__"));
  await handle.promise;
  await Promise.resolve(); // allow persistence save

  const state = await store.load();
  assert.ok(Object.prototype.hasOwnProperty.call(state, "__proto__"));
  assert.equal(state["__proto__"]?.policy?.type, "interval");
  assert.equal(typeof state["__proto__"]?.lastRunAtMs, "number");
  assert.equal(({}).polluted, undefined);

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

test("cron: parse supports wildcards, lists, ranges, and steps", () => {
  const schedule = parseCronExpression("*/20 0,12 1-5 1,6 0-6");
  assert.deepEqual(schedule.minutes, [0, 20, 40]);
  assert.deepEqual(schedule.hours, [0, 12]);
  assert.deepEqual(schedule.daysOfMonth, [1, 2, 3, 4, 5]);
  assert.deepEqual(schedule.months, [1, 6]);
  assert.deepEqual(schedule.daysOfWeek, [0, 1, 2, 3, 4, 5, 6]);
});

test("cron: nextCronRun returns the next matching minute (UTC)", () => {
  const schedule = parseCronExpression("*/15 * * * *");
  assert.equal(nextCronRun(schedule, 0, "utc"), 15 * 60 * 1000);
});

test("RefreshManager: cron policy schedules refreshes and reschedules on register/unregister", async () => {
  const engine = new ControlledEngine();
  const timers = new FakeTimers();
  const manager = new RefreshManager({
    engine,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    timezone: "utc",
  });

  manager.registerQuery(makeQuery("q_cron", { type: "cron", cron: "*/5 * * * *" }));
  // Update policy before the first tick to ensure we reschedule.
  manager.registerQuery(makeQuery("q_cron", { type: "cron", cron: "*/10 * * * *" }));

  const completed = new Promise((resolve) => {
    manager.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_cron") resolve(undefined);
    });
  });

  timers.advance(5 * 60 * 1000);
  assert.equal(engine.calls.length, 0, "old cron schedule should be cleared when policy changes");

  timers.advance(5 * 60 * 1000);
  assert.equal(engine.calls.length, 1);
  engine.calls[0].deferred.resolve(makeResult("q_cron"));
  await completed;

  manager.unregisterQuery("q_cron");
  timers.advance(10 * 60 * 1000);
  assert.equal(engine.calls.length, 1, "unregistered cron query should not run again");

  manager.dispose();
});

test("RefreshManager: cron tick dedupes while a refresh is running", async () => {
  const engine = new ControlledEngine();
  const timers = new FakeTimers();
  const manager = new RefreshManager({
    engine,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    timezone: "utc",
  });

  manager.registerQuery(makeQuery("q_dedupe", { type: "cron", cron: "* * * * *" }));

  const completed = new Promise((resolve) => {
    manager.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_dedupe") resolve(undefined);
    });
  });

  timers.advance(60 * 1000);
  assert.equal(engine.calls.length, 1);

  // Next cron tick fires while the refresh is still running; it should dedupe.
  timers.advance(60 * 1000);
  assert.equal(engine.calls.length, 1);

  engine.calls[0].deferred.resolve(makeResult("q_dedupe"));
  await completed;

  timers.advance(60 * 1000);
  assert.equal(engine.calls.length, 2, "a future tick should start a new refresh once the previous completes");
  engine.calls[1].deferred.resolve(makeResult("q_dedupe"));
  await engine.calls[1].deferred.promise;

  manager.dispose();
});

test("RefreshManager: state store restores interval schedules across instances", async () => {
  const store = new InMemoryRefreshStateStore();
  const timers = new FakeTimers();

  const engine1 = new ControlledEngine();
  const manager1 = new RefreshManager({
    engine: engine1,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    stateStore: store,
  });

  manager1.registerQuery(makeQuery("q_persist", { type: "interval", intervalMs: 10 }));
  await manager1.ready;

  const completed1 = new Promise((resolve) => {
    manager1.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_persist") resolve(undefined);
    });
  });

  timers.advance(10);
  assert.equal(engine1.calls.length, 1);
  engine1.calls[0].deferred.resolve(makeResult("q_persist"));
  await completed1;
  await Promise.resolve(); // allow persistence save

  timers.advance(5); // now = 15
  manager1.dispose();

  const engine2 = new ControlledEngine();
  const manager2 = new RefreshManager({
    engine: engine2,
    concurrency: 1,
    timers: { setTimeout: (...args) => timers.setTimeout(...args), clearTimeout: (id) => timers.clearTimeout(id) },
    now: () => timers.now,
    stateStore: store,
  });

  const query2 = makeQuery("q_persist");
  // Simulate a host that didn't rehydrate refresh settings into the query model.
  // The manager should fall back to the persisted policy.
  delete query2.refreshPolicy;
  manager2.registerQuery(query2);
  await manager2.ready;

  const completed2 = new Promise((resolve) => {
    manager2.onEvent((evt) => {
      if (evt.type === "completed" && evt.job.queryId === "q_persist") resolve(undefined);
    });
  });

  timers.advance(4);
  assert.equal(engine2.calls.length, 0);
  timers.advance(1);
  assert.equal(engine2.calls.length, 1, "next interval run should be scheduled relative to persisted lastRunAtMs");
  engine2.calls[0].deferred.resolve(makeResult("q_persist"));
  await completed2;

  manager2.dispose();
});
