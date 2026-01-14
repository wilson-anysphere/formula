/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as engineModule from "@formula/engine";

import { SpreadsheetApp } from "../spreadsheetApp";

vi.mock("@formula/engine", () => ({
  createEngineClient: vi.fn(),
  engineHydrateFromDocument: vi.fn(async () => []),
  engineApplyDocumentChange: vi.fn(async () => []),
}));

let priorGridMode: string | undefined;

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

function createRoot(): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  (root as any).setPointerCapture ??= () => {};
  (root as any).releasePointerCapture ??= () => {};
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp WASM init teardown", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });
    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 1 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    // Ensure initWasmEngine does not early-return in jsdom.
    const workerCtor = class {};
    Object.defineProperty(globalThis, "Worker", { configurable: true, value: workerCtor });
    Object.defineProperty(window, "Worker", { configurable: true, value: workerCtor });
  });

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("terminates a still-initializing wasm engine client on dispose()", async () => {
    const createEngineClient = (engineModule as any).createEngineClient as ReturnType<typeof vi.fn>;
    const engineHydrateFromDocument = (engineModule as any).engineHydrateFromDocument as ReturnType<typeof vi.fn>;

    expect(typeof (globalThis as any).Worker).toBe("function");
    expect(typeof (window as any).Worker).toBe("function");

    let resolveInit: (() => void) | null = null;
    const initPromise = new Promise<void>((resolve) => {
      resolveInit = resolve;
    });
    const terminate = vi.fn();
    const mockEngine = {
      init: vi.fn(() => initPromise),
      terminate,
    };
    createEngineClient.mockReturnValue(mockEngine);

    const initSpy = vi.spyOn(SpreadsheetApp.prototype as any, "initWasmEngine");

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(initSpy).toHaveBeenCalled();

    // Allow the `wasmSyncPromise.then(...)` init chain to start and call `createEngineClient`.
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(createEngineClient).toHaveBeenCalled();
    expect((app as any).wasmEngine).toBeNull();
    expect((app as any).wasmEngineInit).toBe(mockEngine);

    app.dispose();

    expect(terminate).toHaveBeenCalled();
    expect((app as any).wasmEngine).toBeNull();
    expect((app as any).wasmEngineInit).toBeNull();

    // Ensure teardown does not hang waiting for an init() promise that never resolves.
    await expect(
      Promise.race([
        app.whenIdle(),
        new Promise<void>((_, reject) => setTimeout(() => reject(new Error("Timed out waiting for whenIdle")), 2000)),
      ]),
    ).resolves.toBeUndefined();

    // Even if the init promise resolves later, it should not reattach the engine.
    resolveInit?.();
    await Promise.resolve();
    await Promise.resolve();

    expect(engineHydrateFromDocument).not.toHaveBeenCalled();
    expect((app as any).wasmEngine).toBeNull();
    expect((app as any).wasmEngineInit).toBeNull();

    root.remove();
  });
});
