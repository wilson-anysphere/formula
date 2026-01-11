import { afterEach, describe, expect, it } from "vitest";

import { PyodideRuntime } from "@formula/python-runtime/pyodide";

type GlobalDescriptor = PropertyDescriptor | undefined;

function captureGlobal(name: string): GlobalDescriptor {
  return Object.getOwnPropertyDescriptor(globalThis, name);
}

function restoreGlobal(name: string, descriptor: GlobalDescriptor) {
  if (descriptor) {
    Object.defineProperty(globalThis, name, descriptor);
    return;
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  delete (globalThis as any)[name];
}

class FakeWorker {
  private listeners = new Map<string, Set<(event: any) => void>>();
  terminated = false;

  constructor(_url: unknown, _options?: unknown) {}

  addEventListener(type: string, listener: (event: any) => void) {
    const set = this.listeners.get(type) ?? new Set();
    set.add(listener);
    this.listeners.set(type, set);
  }

  removeEventListener(type: string, listener: (event: any) => void) {
    this.listeners.get(type)?.delete(listener);
  }

  postMessage(msg: any) {
    if (msg?.type === "init") {
      this.emit("message", { data: { type: "ready" } });
    }
  }

  terminate() {
    this.terminated = true;
  }

  private emit(type: string, event: any) {
    for (const listener of this.listeners.get(type) ?? []) {
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
    setStdout() {},
    setStderr() {},
    setInterruptBuffer() {},
    _module: { wasmMemory: { buffer: new ArrayBuffer(0) } },
  };
}

describe("PyodideRuntime mode selection", () => {
  const originals = new Map<string, GlobalDescriptor>();

  function setGlobal(name: string, value: any) {
    if (!originals.has(name)) originals.set(name, captureGlobal(name));
    Object.defineProperty(globalThis, name, { value, writable: true, configurable: true });
  }

  afterEach(() => {
    for (const [name, descriptor] of originals) {
      restoreGlobal(name, descriptor);
    }
    originals.clear();
  });

  it("auto selects worker when crossOriginIsolated + SharedArrayBuffer are available", async () => {
    setGlobal("crossOriginIsolated", true);
    setGlobal("Worker", FakeWorker);

    const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "auto" });
    expect(runtime.getBackendMode()).toBe("worker");

    await runtime.initialize({ api: {}, formulaFiles: {} });
    expect(runtime.backendMode).toBe("worker");
    expect(runtime.initialized).toBe(true);
    runtime.destroy();
  });

  it("auto selects mainThread when SharedArrayBuffer/crossOriginIsolated are unavailable", async () => {
    setGlobal("crossOriginIsolated", false);
    setGlobal("Worker", FakeWorker);
    setGlobal("loadPyodide", async () => createStubPyodide());

    const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "auto" });
    expect(runtime.getBackendMode()).toBe("mainThread");

    await runtime.initialize({ api: {}, formulaFiles: {} });
    expect(runtime.backendMode).toBe("mainThread");
    expect(runtime.initialized).toBe(true);
    runtime.destroy();
  });

  it("throws a deterministic error when neither backend can initialize", async () => {
    setGlobal("crossOriginIsolated", false);
    setGlobal("Worker", undefined);
    setGlobal("SharedArrayBuffer", undefined);
    setGlobal("loadPyodide", undefined);
    setGlobal("document", undefined);

    const runtime = new PyodideRuntime({ api: {}, formulaFiles: {}, mode: "auto" });

    await expect(runtime.initialize({ api: {}, formulaFiles: {} })).rejects.toMatchObject({
      message: "PyodideRuntime mainThread mode requires a DOM (document) or a preloaded globalThis.loadPyodide",
    });
  });
});

