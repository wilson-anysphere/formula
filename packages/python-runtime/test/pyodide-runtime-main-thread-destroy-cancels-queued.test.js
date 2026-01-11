import test from "node:test";
import assert from "node:assert/strict";

import { PyodideRuntime } from "../src/pyodide-runtime.js";

function createSharedStubPyodide() {
  /** @type {string[]} */
  const calls = [];

  /** @type {(() => void) | null} */
  let unblockOne = null;
  const onePromise = new Promise((resolve) => {
    unblockOne = resolve;
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
        return await onePromise;
      }

      if (code.includes('print("two")')) {
        calls.push("two");
        return null;
      }

      // bootstrap/install snippets
      return null;
    },
    __unblockOne() {
      unblockOne?.();
      unblockOne = null;
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

test("queued main-thread execute() rejects if the runtime is destroyed before it starts", async () => {
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
      const runtime1 = new PyodideRuntime({ api: { id: 1 }, formulaFiles: {}, mode: "mainThread" });
      const runtime2 = new PyodideRuntime({ api: { id: 2 }, formulaFiles: {}, mode: "mainThread" });

      await runtime1.initialize();
      await runtime2.initialize();

      pyodide.__calls.length = 0;

      const p1 = runtime1.execute('print("one")\n');
      const p2 = runtime2.execute('print("two")\n');

      // Let the first execution start (and block).
      await new Promise((resolve) => setTimeout(resolve, 0));
      assert.deepEqual(pyodide.__calls, ["sandbox", "one"]);

      runtime2.destroy();

      pyodide.__unblockOne();
      await p1;

      await assert.rejects(p2, { message: "PyodideRuntime was destroyed" });
      assert.deepEqual(pyodide.__calls, ["sandbox", "one"]);

      runtime1.destroy();
    },
  );
});

