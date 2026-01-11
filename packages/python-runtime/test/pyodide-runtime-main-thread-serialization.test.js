import test from "node:test";
import assert from "node:assert/strict";

import { PyodideRuntime } from "../src/pyodide-runtime.js";

function createSharedStubPyodide() {
  /** @type {string[]} */
  const calls = [];

  /** @type {(() => void) | null} */
  let unblockFirst = null;
  const firstUserPromise = new Promise((resolve) => {
    unblockFirst = resolve;
  });

  const runtime = {
    FS: { mkdirTree() {}, writeFile() {} },
    registerJsModule() {},
    setInterruptBuffer() {},
    setStdout() {},
    setStderr() {},
    _module: { wasmMemory: { buffer: new ArrayBuffer(0) } },
    async runPythonAsync(code) {
      if (typeof code === "string" && code.includes("apply_sandbox")) {
        calls.push("sandbox");
        return null;
      }

      if (code.includes('print("one")')) {
        calls.push("one");
        return await firstUserPromise;
      }

      if (code.includes('print("two")')) {
        calls.push("two");
        return null;
      }

      // bootstrap/install snippets
      return null;
    },
    __unblockFirst() {
      unblockFirst?.();
      unblockFirst = null;
    },
    __calls: calls,
  };

  return runtime;
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

test("main-thread Pyodide executions are serialized across runtime instances (cached Pyodide)", async () => {
  const pyodide = createSharedStubPyodide();

  await withStubbedGlobals(
    {
      crossOriginIsolated: false,
      SharedArrayBuffer: undefined,
      Worker: undefined,
      loadPyodide: async () => pyodide,
      document: undefined,
    },
    async () => {
      const runtime1 = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "mainThread" });
      const runtime2 = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "mainThread" });

      await runtime1.initialize();
      await runtime2.initialize();

      // Drop bootstrap calls (we only care about execute ordering).
      pyodide.__calls.length = 0;

      const p1 = runtime1.execute('print("one")\n');
      const p2 = runtime2.execute('print("two")\n');

      // Give the event loop a tick; runtime2 should not start while runtime1 is blocked.
      await new Promise((resolve) => setTimeout(resolve, 0));
      assert.deepEqual(pyodide.__calls, ["sandbox", "one"]);

      pyodide.__unblockFirst();
      await p1;
      await p2;

      assert.deepEqual(pyodide.__calls, ["sandbox", "one", "sandbox", "two"]);

      runtime1.destroy();
      runtime2.destroy();
    },
  );
});

