import test from "node:test";
import assert from "node:assert/strict";

import { PyodideRuntime } from "../src/pyodide-runtime.js";

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

test("PyodideRuntime.destroy clears the main-thread bridge API from the Pyodide instance", async () => {
  const pyodide = createStubPyodide();
  const api = { get_active_sheet_id() {} };

  await withStubbedGlobals(
    {
      crossOriginIsolated: false,
      SharedArrayBuffer: undefined,
      Worker: undefined,
      loadPyodide: async () => pyodide,
      document: undefined,
    },
    async () => {
      const runtime = new PyodideRuntime({ api, formulaFiles: {}, mode: "mainThread" });
      await runtime.initialize();

      assert.equal(pyodide.__formulaPyodideBridgeApi, api);

      runtime.destroy();
      assert.equal(pyodide.__formulaPyodideBridgeApi, null);
    },
  );
});

