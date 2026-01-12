import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DesktopExtensionHostManager } from "./extensionHostManager.js";

class FakeWorker {
  private readonly listeners = new Map<string, Set<(evt: any) => void>>();
  private terminated = false;

  addEventListener(type: string, listener: (evt: any) => void) {
    const key = String(type);
    const set = this.listeners.get(key) ?? new Set();
    set.add(listener);
    this.listeners.set(key, set);
  }

  removeEventListener(type: string, listener: (evt: any) => void) {
    const key = String(type);
    const set = this.listeners.get(key);
    if (!set) return;
    set.delete(listener);
    if (set.size === 0) this.listeners.delete(key);
  }

  postMessage(_message: unknown) {
    // Host -> worker messages are ignored for this unit test.
  }

  terminate() {
    this.terminated = true;
  }

  emitMessage(message: unknown) {
    if (this.terminated) return;
    const set = this.listeners.get("message");
    if (!set) return;
    for (const listener of [...set]) {
      listener({ data: message });
    }
  }
}

describe("DesktopExtensionHostManager clipboard wiring", () => {
  const workerInstances: FakeWorker[] = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const PrevWorker = (globalThis as any).Worker;

  beforeEach(() => {
    workerInstances.length = 0;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Worker = class extends FakeWorker {
      constructor(_url: unknown, _options?: unknown) {
        super();
        workerInstances.push(this);
      }
    };
  });

  afterEach(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Worker = PrevWorker;
  });

  it("routes clipboard.writeText API calls to the injected clipboardApi adapter", async () => {
    let resolveWrite!: () => void;
    const writeCalled = new Promise<void>((resolve) => {
      resolveWrite = resolve;
    });

    const clipboardApi = {
      readText: vi.fn(async () => ""),
      writeText: vi.fn(async (_text: string) => {
        resolveWrite();
      }),
    };

    const manager = new DesktopExtensionHostManager({
      engineVersion: "1.0.0",
      spreadsheetApi: {},
      uiApi: {},
      permissionPrompt: async () => true,
      clipboardApi,
    });

    const extensionId = "test.clipboard";
    await manager.host.loadExtension({
      extensionId,
      extensionPath: "memory://clipboard/",
      manifest: {
        name: "clipboard",
        publisher: "test",
        version: "1.0.0",
        engines: { formula: "^1.0.0" },
        main: "./dist/extension.mjs",
        activationEvents: [],
        permissions: ["clipboard"],
        contributes: {},
      },
      mainUrl: "memory://clipboard/main.mjs",
    });

    const worker = workerInstances[0];
    expect(worker).toBeTruthy();

    worker.emitMessage({
      type: "api_call",
      id: "req-1",
      namespace: "clipboard",
      method: "writeText",
      args: ["hello from extension"],
    });

    await writeCalled;
    expect(clipboardApi.writeText).toHaveBeenCalledWith("hello from extension");

    await manager.host.dispose();
  });

  it("invokes clipboardWriteGuard before delegating to clipboardApi.writeText", async () => {
    const steps: string[] = [];
    let resolveGuard!: () => void;
    const guardCalled = new Promise<void>((resolve) => {
      resolveGuard = resolve;
    });

    const clipboardWriteGuard = vi.fn(async () => {
      steps.push("guard");
      resolveGuard();
    });

    let resolveWrite!: () => void;
    const writeCalled = new Promise<void>((resolve) => {
      resolveWrite = resolve;
    });

    const clipboardApi = {
      readText: vi.fn(async () => ""),
      writeText: vi.fn(async () => {
        steps.push("write");
        resolveWrite();
      }),
    };

    const manager = new DesktopExtensionHostManager({
      engineVersion: "1.0.0",
      spreadsheetApi: {},
      uiApi: {},
      permissionPrompt: async () => true,
      clipboardApi,
      clipboardWriteGuard,
    });

    const extensionId = "test.clipboard-guard";
    await manager.host.loadExtension({
      extensionId,
      extensionPath: "memory://clipboard-guard/",
      manifest: {
        name: "clipboard-guard",
        publisher: "test",
        version: "1.0.0",
        engines: { formula: "^1.0.0" },
        main: "./dist/extension.mjs",
        activationEvents: [],
        permissions: ["clipboard"],
        contributes: {},
      },
      mainUrl: "memory://clipboard-guard/main.mjs",
    });

    const worker = workerInstances[0];
    expect(worker).toBeTruthy();

    worker.emitMessage({
      type: "api_call",
      id: "req-guard",
      namespace: "clipboard",
      method: "writeText",
      args: ["guarded"],
    });

    await guardCalled;
    await writeCalled;
    expect(clipboardWriteGuard).toHaveBeenCalledWith({ extensionId, taintedRanges: [] });
    expect(clipboardApi.writeText).toHaveBeenCalledWith("guarded");
    expect(steps).toEqual(["guard", "write"]);

    await manager.host.dispose();
  });
});
