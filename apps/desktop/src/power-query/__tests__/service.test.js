import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../../../../packages/power-query/src/table.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryService } from "../service.ts";

function createMemoryStorage() {
  const map = new Map();
  return {
    getItem(key) {
      const value = map.get(String(key));
      return value === undefined ? null : value;
    },
    setItem(key, value) {
      map.set(String(key), String(value));
    },
    removeItem(key) {
      map.delete(String(key));
    },
    clear() {
      map.clear();
    },
  };
}

function createFakeTimers() {
  let nowMs = 0;
  let nextId = 1;
  /** @type {Map<number, { id: number, time: number, order: number, fn: () => void }>} */
  const tasks = new Map();

  /** @type {typeof setTimeout} */
  const setTimeoutFn = (fn, delay, ...args) => {
    const id = nextId++;
    const ms = Number.isFinite(delay) ? Number(delay) : 0;
    const time = nowMs + Math.max(0, ms);
    tasks.set(id, { id, time, order: id, fn: () => fn(...args) });
    return id;
  };

  /** @type {typeof clearTimeout} */
  const clearTimeoutFn = (id) => {
    tasks.delete(Number(id));
  };

  function tick(ms) {
    const target = nowMs + ms;
    while (true) {
      let next = null;
      for (const task of tasks.values()) {
        if (task.time > target) continue;
        if (!next || task.time < next.time || (task.time === next.time && task.order < next.order)) {
          next = task;
        }
      }
      if (!next) break;
      tasks.delete(next.id);
      nowMs = next.time;
      next.fn();
    }
    nowMs = target;
  }

  return {
    timers: { setTimeout: setTimeoutFn, clearTimeout: clearTimeoutFn },
    tick,
    now: () => nowMs,
  };
}

async function flushMicrotasks(times = 10) {
  let remaining = times;
  while (remaining > 0) {
    remaining -= 1;
    await Promise.resolve();
  }
}

test("DesktopPowerQueryService schedules interval refreshes without requiring the Query Editor UI", async () => {
  const originalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const storage = createMemoryStorage();
  Object.defineProperty(globalThis, "localStorage", { value: storage, configurable: true });

  try {
    const fake = createFakeTimers();
    let refreshCalls = 0;

    const table = DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true });
    const engine = {
      async executeQueryWithMeta(query) {
        refreshCalls += 1;
        return { table, meta: { queryId: query.id } };
      },
    };

    const service = new DesktopPowerQueryService({
      workbookId: "wb_interval",
      document: new DocumentController({ engine: new MockEngine() }),
      engine,
      getContext: () => ({}),
      concurrency: 1,
      batchSize: 1,
      refresh: { timers: fake.timers, now: fake.now },
    });

    const query = {
      id: "q_interval",
      name: "Interval",
      source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "interval", intervalMs: 1_000 },
    };

    const completed = new Promise((resolve) => {
      const unsub = service.onEvent((evt) => {
        if (evt?.type === "completed" && evt?.job?.queryId === query.id) {
          unsub();
          resolve(evt);
        }
      });
    });

    service.registerQuery(query);
    await flushMicrotasks();

    fake.tick(1_000);
    await completed;

    assert.equal(refreshCalls, 1);

    service.dispose();

    fake.tick(5_000);
    await flushMicrotasks();
    assert.equal(refreshCalls, 1, "expected dispose() to cancel scheduled refresh timers");
  } finally {
    if (originalStorageDescriptor) {
      Object.defineProperty(globalThis, "localStorage", originalStorageDescriptor);
    } else {
      delete globalThis.localStorage;
    }
  }
});

test("DesktopPowerQueryService persists query definitions across instances", async () => {
  const originalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const storage = createMemoryStorage();
  Object.defineProperty(globalThis, "localStorage", { value: storage, configurable: true });

  try {
    const engine = {
      async executeQueryWithMeta() {
        return { table: DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true }), meta: {} };
      },
    };

    const workbookId = "wb_persist";
    const query = {
      id: "q_persist",
      name: "Persisted",
      source: { type: "range", range: { values: [["X"], [1]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "manual" },
    };

    const first = new DesktopPowerQueryService({
      workbookId,
      document: new DocumentController({ engine: new MockEngine() }),
      engine,
      getContext: () => ({}),
    });
    first.registerQuery(query);
    first.dispose();

    const second = new DesktopPowerQueryService({
      workbookId,
      document: new DocumentController({ engine: new MockEngine() }),
      engine,
      getContext: () => ({}),
    });

    const reloaded = second.getQueries();
    assert.equal(reloaded.length, 1);
    assert.equal(reloaded[0].id, query.id);
    assert.equal(reloaded[0].name, query.name);
    second.dispose();
  } finally {
    if (originalStorageDescriptor) {
      Object.defineProperty(globalThis, "localStorage", originalStorageDescriptor);
    } else {
      delete globalThis.localStorage;
    }
  }
});
