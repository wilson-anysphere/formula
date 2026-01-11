import test from "node:test";
import assert from "node:assert/strict";

import { PyodideRuntime } from "../src/pyodide-runtime.js";

class FakeWorker {
  addEventListener() {}
  removeEventListener() {}
  postMessage() {}
  terminate() {}
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

test("PyodideRuntime worker mode rejects when SharedArrayBuffer/crossOriginIsolated are unavailable", async () => {
  await withStubbedGlobals(
    {
      Worker: FakeWorker,
      SharedArrayBuffer: undefined,
      crossOriginIsolated: false,
    },
    async () => {
      const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "worker" });
      await assert.rejects(runtime.initialize(), {
        message:
          "PyodideRuntime worker mode requires crossOriginIsolated + SharedArrayBuffer (COOP/COEP). Use mode: 'mainThread' to run without SharedArrayBuffer (UI may freeze).",
      });
    },
  );
});
