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

