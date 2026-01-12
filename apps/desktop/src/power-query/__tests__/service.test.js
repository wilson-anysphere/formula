import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "@formula/power-query";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { DesktopPowerQueryService, saveQueriesToStorage } from "../service.ts";

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

test("DesktopPowerQueryService prefers workbook-backed queries over localStorage", async () => {
  const originalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const storage = createMemoryStorage();
  Object.defineProperty(globalThis, "localStorage", { value: storage, configurable: true });

  try {
    const workbookId = "wb_workbook_part";

    const localQuery = {
      id: "q_local",
      name: "LocalStorage",
      source: { type: "range", range: { values: [["A"], [1]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "manual" },
    };
    saveQueriesToStorage(workbookId, [localQuery]);

    const workbookQuery = {
      id: "q_workbook",
      name: "WorkbookPart",
      source: { type: "range", range: { values: [["B"], [2]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "manual" },
    };
    const xml = `<FormulaPowerQuery version="1"><![CDATA[${JSON.stringify({ queries: [workbookQuery] })}]]></FormulaPowerQuery>`;

    /** @type {{ cmd: string, args?: any }[]} */
    const calls = [];
    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      value: {
        core: {
          invoke: async (cmd, args) => {
            calls.push({ cmd, args });
            if (cmd === "power_query_state_get") return xml;
            if (cmd === "power_query_state_set") return null;
            throw new Error(`Unexpected invoke(${cmd})`);
          },
        },
      },
    });

    const service = new DesktopPowerQueryService({
      workbookId,
      document: new DocumentController({ engine: new MockEngine() }),
      engine: { async executeQueryWithMeta() { return { table: DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true }), meta: {} }; } },
      getContext: () => ({}),
    });
    await service.ready;

    const queries = service.getQueries();
    assert.equal(queries.length, 1);
    assert.equal(queries[0].id, workbookQuery.id);

    assert.ok(calls.some((c) => c.cmd === "power_query_state_get"), "expected service to call power_query_state_get");
    service.dispose();
  } finally {
    if (originalStorageDescriptor) {
      Object.defineProperty(globalThis, "localStorage", originalStorageDescriptor);
    } else {
      delete globalThis.localStorage;
    }
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
    } else {
      delete globalThis.__TAURI__;
    }
  }
});

test("DesktopPowerQueryService persists queries to the workbook part via Tauri", async () => {
  const originalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const storage = createMemoryStorage();
  Object.defineProperty(globalThis, "localStorage", { value: storage, configurable: true });

  try {
    /** @type {{ cmd: string, args?: any }[]} */
    const calls = [];
    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      value: {
        core: {
          invoke: async (cmd, args) => {
            calls.push({ cmd, args });
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;
            throw new Error(`Unexpected invoke(${cmd})`);
          },
        },
      },
    });

    const doc = new DocumentController({ engine: new MockEngine() });
    const service = new DesktopPowerQueryService({
      workbookId: "wb_persist_workbook_part",
      document: doc,
      engine: { async executeQueryWithMeta() { return { table: DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true }), meta: {} }; } },
      getContext: () => ({}),
    });
    await service.ready;

    service.registerQuery({
      id: "q1",
      name: "Persisted",
      source: { type: "range", range: { values: [["A"], [1]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "manual" },
      credentials: { token: "should-not-be-persisted" },
    });
    await flushMicrotasks();

    const setCalls = calls.filter((c) => c.cmd === "power_query_state_set");
    assert.ok(setCalls.length >= 1, "expected persist to call power_query_state_set");
    const last = setCalls[setCalls.length - 1];
    assert.ok(typeof last.args?.xml === "string" && last.args.xml.includes("<FormulaPowerQuery"), "expected XML payload");
    assert.ok(!last.args.xml.includes("should-not-be-persisted"), "expected credentials to be redacted");
    assert.equal(doc.isDirty, true, "expected persisting queries to mark the document dirty under Tauri");

    service.dispose();
  } finally {
    if (originalStorageDescriptor) {
      Object.defineProperty(globalThis, "localStorage", originalStorageDescriptor);
    } else {
      delete globalThis.localStorage;
    }
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
    } else {
      delete globalThis.__TAURI__;
    }
  }
});

test("DesktopPowerQueryService escapes CDATA terminators when persisting workbook XML", async () => {
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");

  try {
    let persistedXml = null;
    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      value: {
        core: {
          invoke: async (cmd, args) => {
            if (cmd === "power_query_state_get") return persistedXml;
            if (cmd === "power_query_state_set") {
              persistedXml = args?.xml ?? null;
              return null;
            }
            throw new Error(`Unexpected invoke(${cmd})`);
          },
        },
      },
    });

    const queryWithCdataTerminator = {
      id: "q_cdata",
      name: "Weird]]>Name",
      source: { type: "range", range: { values: [["A"], [1]], hasHeaders: true } },
      steps: [],
      refreshPolicy: { type: "manual" },
    };

    const doc1 = new DocumentController({ engine: new MockEngine() });
    const first = new DesktopPowerQueryService({
      workbookId: "wb_cdata",
      document: doc1,
      engine: {
        async executeQueryWithMeta() {
          return { table: DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true }), meta: {} };
        },
      },
      getContext: () => ({}),
    });
    await first.ready;
    first.registerQuery(queryWithCdataTerminator);
    await flushMicrotasks();

    assert.ok(typeof persistedXml === "string" && persistedXml.includes("<![CDATA["), "expected a persisted XML payload");
    assert.ok(
      persistedXml.includes("Weird]]\\u003eName"),
      "expected JSON to escape ']]>' as a unicode escape sequence inside CDATA",
    );
    assert.ok(!persistedXml.includes("Weird]]>Name"), "expected raw ']]>' substring to not appear inside CDATA content");
    first.dispose();

    const doc2 = new DocumentController({ engine: new MockEngine() });
    const second = new DesktopPowerQueryService({
      workbookId: "wb_cdata",
      document: doc2,
      engine: {
        async executeQueryWithMeta() {
          return { table: DataTable.fromGrid([["A"], [1]], { hasHeaders: true, inferTypes: true }), meta: {} };
        },
      },
      getContext: () => ({}),
    });
    await second.ready;
    const loaded = second.getQueries();
    assert.equal(loaded.length, 1);
    assert.equal(loaded[0].name, queryWithCdataTerminator.name);
    second.dispose();
  } finally {
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
    } else {
      delete globalThis.__TAURI__;
    }
  }
});
