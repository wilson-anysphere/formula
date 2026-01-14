import assert from "node:assert/strict";
import test from "node:test";

import { ScriptRuntime } from "../src/web.js";

function restoreProperty(target, prop, descriptor) {
  if (descriptor) {
    Object.defineProperty(target, prop, descriptor);
    return;
  }
  // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
  delete target[prop];
}

test("ScriptRuntime (web): ui.alert prefers Tauri plugin dialog even when __TAURI__.dialog getter throws", async (t) => {
  const originalTauri = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const originalAlert = Object.getOwnPropertyDescriptor(globalThis, "alert");

  t.after(() => {
    restoreProperty(globalThis, "__TAURI__", originalTauri);
    restoreProperty(globalThis, "alert", originalAlert);
  });

  let alertCalled = false;
  globalThis.alert = () => {
    alertCalled = true;
  };

  const dialog = {};
  /** @type {any} */
  let dialogThis = null;
  /** @type {any} */
  dialog.message = function (msg) {
    dialogThis = this;
    assert.equal(msg, "hello");
  };

  const tauri = { plugin: { dialog } };
  Object.defineProperty(tauri, "dialog", {
    configurable: true,
    get() {
      throw new Error("blocked dialog access");
    },
  });
  Object.defineProperty(globalThis, "__TAURI__", { configurable: true, value: tauri, writable: true });

  const runtime = new ScriptRuntime({});
  await runtime.handleRpc("ui.alert", { message: "hello" });

  assert.equal(alertCalled, false);
  assert.equal(dialogThis, dialog);
});

test("ScriptRuntime (web): ui.confirm prefers Tauri plugin dialog even when __TAURI__.dialog getter throws", async (t) => {
  const originalTauri = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const originalConfirm = Object.getOwnPropertyDescriptor(globalThis, "confirm");

  t.after(() => {
    restoreProperty(globalThis, "__TAURI__", originalTauri);
    restoreProperty(globalThis, "confirm", originalConfirm);
  });

  let confirmCalled = false;
  globalThis.confirm = () => {
    confirmCalled = true;
    return false;
  };

  const dialog = {};
  /** @type {any} */
  let dialogThis = null;
  /** @type {any} */
  dialog.confirm = function (msg) {
    dialogThis = this;
    assert.equal(msg, "are you sure?");
    return true;
  };

  const tauri = { plugins: { dialog } };
  Object.defineProperty(tauri, "dialog", {
    configurable: true,
    get() {
      throw new Error("blocked dialog access");
    },
  });
  Object.defineProperty(globalThis, "__TAURI__", { configurable: true, value: tauri, writable: true });

  const runtime = new ScriptRuntime({});
  const result = await runtime.handleRpc("ui.confirm", { message: "are you sure?" });

  assert.equal(confirmCalled, false);
  assert.equal(dialogThis, dialog);
  assert.equal(result, true);
});

