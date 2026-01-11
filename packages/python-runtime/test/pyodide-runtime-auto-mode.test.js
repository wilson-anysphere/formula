import test from "node:test";
import assert from "node:assert/strict";

import { PyodideRuntime } from "../src/pyodide-runtime.js";

class FakeWorker {
  /** @type {Map<string, Set<(event: any) => void>>} */
  #listeners = new Map();

  constructor(_url, _options) {}

  addEventListener(type, listener) {
    const set = this.#listeners.get(type) ?? new Set();
    set.add(listener);
    this.#listeners.set(type, set);
  }

  removeEventListener(type, listener) {
    this.#listeners.get(type)?.delete(listener);
  }

  postMessage(msg) {
    if (msg?.type === "init") {
      queueMicrotask(() => this.#emit("message", { data: { type: "ready" } }));
    }
  }

  terminate() {}

  #emit(type, event) {
    for (const listener of this.#listeners.get(type) ?? []) {
      listener(event);
    }
  }
}

function createStubPyodide() {
  return {
    FS: { mkdirTree() {}, writeFile() {} },
    registerJsModule() {},
    async runPythonAsync() {
      return null;
    },
    setInterruptBuffer() {},
  };
}

function withStubbedGlobals(stubs, fn) {
  const originals = new Map();
  for (const [key, value] of Object.entries(stubs)) {
    originals.set(key, Object.getOwnPropertyDescriptor(globalThis, key));
    Object.defineProperty(globalThis, key, { value, writable: true, configurable: true });
  }
  return Promise.resolve()
    .then(fn)
    .finally(() => {
      for (const [key, descriptor] of originals) {
        if (descriptor) {
          Object.defineProperty(globalThis, key, descriptor);
        } else {
          delete globalThis[key];
        }
      }
    });
}

test("PyodideRuntime auto mode selects worker when SAB + crossOriginIsolated are available", async () => {
  await withStubbedGlobals(
    {
      Worker: FakeWorker,
      crossOriginIsolated: true,
    },
    async () => {
      const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "auto" });
      assert.equal(runtime.getBackendMode(), "worker");
      await runtime.initialize({ api: {}, formulaFiles: {} });
      assert.equal(runtime.backendMode, "worker");
      runtime.destroy();
    },
  );
});

test("PyodideRuntime auto mode selects mainThread when crossOriginIsolated is false", async () => {
  await withStubbedGlobals(
    {
      Worker: FakeWorker,
      crossOriginIsolated: false,
      loadPyodide: async () => createStubPyodide(),
      document: undefined,
    },
    async () => {
      const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "auto" });
      assert.equal(runtime.getBackendMode(), "mainThread");
      await runtime.initialize({ api: {}, formulaFiles: {} });
      assert.equal(runtime.backendMode, "mainThread");
      runtime.destroy();
    },
  );
});

