import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

function createMockPyodide() {
  /** @type {{ batched?: (text: string) => void } | null} */
  let stdoutConfig = null;
  /** @type {{ batched?: (text: string) => void } | null} */
  let stderrConfig = null;

  return {
    setInterruptBuffer() {},
    setStdout(config) {
      stdoutConfig = config;
    },
    setStderr(config) {
      stderrConfig = config;
    },
    registerJsModule() {},
    FS: { mkdirTree() {}, writeFile() {} },
    async runPythonAsync(code) {
      // Sandbox/bootstrap snippets should not emit output.
      if (code.includes("print(")) {
        const label = code.includes("first") ? "first" : "second";
        stdoutConfig?.batched?.(`${label} stdout\n`);
        stderrConfig?.batched?.(`${label} stderr\n`);
      }
      return null;
    },
  };
}

test("pyodide worker includes stdout/stderr fields and resets them per execute", async () => {
  const workerPath = path.join(
    path.dirname(fileURLToPath(import.meta.url)),
    "..",
    "src",
    "pyodide-worker.js",
  );
  const workerSource = await readFile(workerPath, "utf8");

  /** @type {Array<any>} */
  const posted = [];
  /** @type {Map<string, { resolve: (msg: any) => void, reject: (err: any) => void, timer: any }>} */
  const waiters = new Map();

  const runtime = createMockPyodide();
  const self = {
    fetch: async () => ({ ok: true }),
    WebSocket: class StubWebSocket {},
    location: { href: "https://localhost" },
    loadPyodide: async () => runtime,
    postMessage(msg) {
      posted.push(msg);
      const waiter = waiters.get(msg?.requestId);
      if (!waiter) return;
      clearTimeout(waiter.timer);
      waiters.delete(msg.requestId);
      waiter.resolve(msg);
    },
  };

  const context = vm.createContext({
    self,
    console,
    TextEncoder,
    TextDecoder,
    SharedArrayBuffer,
    Atomics,
    URL,
    setTimeout,
    clearTimeout,
    importScripts() {},
  });

  vm.runInContext(workerSource, context, { filename: "pyodide-worker.js" });

  async function exec(requestId, code) {
    return await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error("Timed out waiting for worker response")), 1_000);
      waiters.set(requestId, { resolve, reject, timer });
      self.onmessage({
        data: {
          type: "execute",
          requestId,
          code,
          permissions: { filesystem: "none", network: "none" },
        },
      });
    });
  }

  const first = await exec("first", 'print("first")');
  assert.equal(first.success, true);
  assert.ok(Object.hasOwn(first, "stdout"));
  assert.ok(Object.hasOwn(first, "stderr"));
  assert.match(first.stdout, /first stdout/);
  assert.match(first.stderr, /first stderr/);

  const second = await exec("second", 'print("second")');
  assert.equal(second.success, true);
  assert.ok(Object.hasOwn(second, "stdout"));
  assert.ok(Object.hasOwn(second, "stderr"));
  assert.match(second.stdout, /second stdout/);
  assert.match(second.stderr, /second stderr/);

  // Output should be isolated per execution.
  assert.ok(!second.stdout.includes("first stdout"));
  assert.ok(!second.stderr.includes("first stderr"));
});
