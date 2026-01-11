import assert from "node:assert/strict";
import test from "node:test";

import { createPowerQueryRefreshStateStore } from "../refreshStateStore.ts";

test("createPowerQueryRefreshStateStore prefers Tauri invoke when available", async () => {
  /** @type {Array<{ cmd: string, args: any }>} */
  const calls = [];
  const invoke = async (cmd, args) => {
    calls.push({ cmd, args });
    if (cmd === "power_query_refresh_state_get") {
      return { q1: { policy: { type: "manual" } } };
    }
    return null;
  };

  const previous = globalThis.__TAURI__;
  globalThis.__TAURI__ = { core: { invoke } };

  try {
    const store = createPowerQueryRefreshStateStore({ workbookId: "wb_tauri" });
    const loaded = await store.load();

    assert.equal(calls[0]?.cmd, "power_query_refresh_state_get");
    assert.deepEqual(calls[0]?.args, { workbook_id: "wb_tauri" });
    assert.deepEqual(loaded, { q1: { policy: { type: "manual" } } });

    await store.save({ q2: { policy: { type: "manual" }, lastRunAtMs: 123 } });
    assert.equal(calls[1]?.cmd, "power_query_refresh_state_set");
    assert.deepEqual(calls[1]?.args, { workbook_id: "wb_tauri", state: { q2: { policy: { type: "manual" }, lastRunAtMs: 123 } } });
  } finally {
    if (previous === undefined) {
      delete globalThis.__TAURI__;
    } else {
      globalThis.__TAURI__ = previous;
    }
  }
});

test("createPowerQueryRefreshStateStore namespaces persistence by workbookId for storage backends", async () => {
  const map = new Map();
  const storage = {
    getItem(key) {
      return map.get(key) ?? null;
    },
    setItem(key, value) {
      map.set(key, value);
    },
  };

  const storeA = createPowerQueryRefreshStateStore({ workbookId: "wb_a", storage });
  const storeB = createPowerQueryRefreshStateStore({ workbookId: "wb_b", storage });

  await storeA.save({ q1: { policy: { type: "manual" }, lastRunAtMs: 1 } });
  await storeB.save({ q1: { policy: { type: "manual" }, lastRunAtMs: 2 } });

  assert.ok(map.has("formula.desktop.powerQuery.refreshState:wb_a"));
  assert.ok(map.has("formula.desktop.powerQuery.refreshState:wb_b"));

  assert.equal((await storeA.load()).q1.lastRunAtMs, 1);
  assert.equal((await storeB.load()).q1.lastRunAtMs, 2);
});

test("createPowerQueryRefreshStateStore in-memory fallback is keyed by workbookId", async () => {
  const storeA = createPowerQueryRefreshStateStore({ workbookId: "wb_mem_a", storage: null });
  const storeB = createPowerQueryRefreshStateStore({ workbookId: "wb_mem_b", storage: null });

  await storeA.save({ q1: { policy: { type: "manual" }, lastRunAtMs: 10 } });
  assert.deepEqual(await storeB.load(), {});

  // New instance should see the same in-memory state for the same workbook key.
  const storeA2 = createPowerQueryRefreshStateStore({ workbookId: "wb_mem_a", storage: null });
  assert.equal((await storeA2.load()).q1.lastRunAtMs, 10);
});
