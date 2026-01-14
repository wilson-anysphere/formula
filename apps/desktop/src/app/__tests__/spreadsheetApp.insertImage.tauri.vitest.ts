/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import { SpreadsheetApp } from "../spreadsheetApp";

const mocks = vi.hoisted(() => {
  return {
    pickLocalImageFiles: vi.fn<[], Promise<File[]>>(),
  };
});

vi.mock("../../drawings/pickLocalImageFiles.js", () => ({
  pickLocalImageFiles: mocks.pickLocalImageFiles,
}));

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
  document.body.appendChild(root);
  return root;
}

async function waitFor(condition: () => boolean, timeoutMs: number = 1000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for condition");
}

describe("SpreadsheetApp insert image (Tauri picker)", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
    (globalThis as any).__TAURI__ = undefined;
  });

  beforeEach(() => {
    // Ensure this test suite always runs with real timers.
    // Some other suites use fake timers and (if they fail before cleanup) can leak them,
    // which would cause `waitFor()` (Date.now + setTimeout) to hang indefinitely.
    vi.useRealTimers();

    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      writable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, writable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({})),
    });

    // Trigger the Tauri path in SpreadsheetApp.insertImageFromLocalFile.
    (globalThis as any).__TAURI__ = {
      dialog: { open: vi.fn(async () => null) },
      core: { invoke: vi.fn(async () => null) },
    };

    mocks.pickLocalImageFiles.mockReset();
  });

  it(
    "prefers pickLocalImageFiles when Tauri dialog + invoke are available",
    async () => {
      const bytes = new Uint8Array([1, 2, 3, 4, 5, 6]);
      const file = new File([bytes], "test.png", { type: "image/png" });
      mocks.pickLocalImageFiles.mockResolvedValue([file]);

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    app.activateCell({ row: 3, col: 4 }, { scrollIntoView: false, focus: false });
    app.insertImageFromLocalFile();

    await waitFor(() => mocks.pickLocalImageFiles.mock.calls.length === 1);
    expect(mocks.pickLocalImageFiles).toHaveBeenCalledWith({ multiple: false });

    // The native picker path should not create the hidden <input type=file> element.
    expect(root.querySelector('input[data-testid="insert-image-input"]')).toBeNull();

    await waitFor(() => ((app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? []).length === 1);
    const drawings = (app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? [];
    expect(drawings).toHaveLength(1);
    const obj = drawings[0]!;

    expect(obj?.kind?.type).toBe("image");
    expect(obj?.anchor).toEqual({
      type: "oneCell",
      from: { cell: { row: 3, col: 4 }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(200), cy: pxToEmu(150) },
    });

    const entry = app.getDrawingImages().get(obj.kind.imageId);
    expect(entry).toBeTruthy();
    expect(entry.bytes).toEqual(bytes);

    app.destroy();
    root.remove();
    },
    60_000,
  );
});
