/**
 * @vitest-environment jsdom
 */
 
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { createSheetNameResolverFromIdToNameMap } from "../../sheet/sheetNameResolver.js";

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
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp computed-value cache sheet name mapping", () => {
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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("stores + invalidates computed values under stable sheet ids when engine emits display names", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetIdToName = new Map<string, string>([["Sheet1", "Sheet1"]]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

    const app = new SpreadsheetApp(root, status, { sheetNameResolver });
    const doc = app.getDocument();

    // Rename the sheet so its stable id no longer matches the display name.
    doc.renameSheet("Sheet1", "Budget");
    sheetIdToName.set("Sheet1", "Budget");

    await app.whenIdle();

    // Simulate a WASM engine recalc delta that uses the sheet display name.
    (app as any).applyComputedChanges([{ sheet: "Budget", address: "A1", value: 123 }]);

    const computedValuesByCoord = (app as any).computedValuesByCoord as Map<string, Map<number, unknown>>;
    expect(computedValuesByCoord.get("Sheet1")?.get(0)).toBe(123);
    expect(computedValuesByCoord.has("Budget")).toBe(false);

    // Input edits should invalidate the stable-id cache entry even if the sheet token is a display name.
    (app as any).invalidateComputedValues([{ sheet: "Budget", address: "A1" }]);
    expect(computedValuesByCoord.get("Sheet1")?.has(0)).toBe(false);

    app.destroy();
    root.remove();
  });

  it("trims sheet ids returned by sheetNameResolver.getSheetIdByName", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const sheetNameResolver = {
      getSheetNameById: (id: string) => (id === "Sheet1" ? "Budget" : null),
      getSheetIdByName: (name: string) => (name.trim().toLowerCase() === "budget" ? "  Sheet1  " : null),
    };

    const app = new SpreadsheetApp(root, status, { sheetNameResolver: sheetNameResolver as any });
    expect(app.getSheetIdByName("Budget")).toBe("Sheet1");

    app.destroy();
    root.remove();
  });

  it("resolves DocumentController sheet meta display names even when sheetNameResolver is absent", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();

    // Rename the sheet so its stable id no longer matches the display name.
    // No `sheetNameResolver` is provided, so resolution must fall back to DocumentController sheet meta.
    doc.renameSheet("Sheet1", "Budget");

    await app.whenIdle();

    (app as any).applyComputedChanges([{ sheet: "Budget", address: "A1", value: 123 }]);

    const computedValuesByCoord = (app as any).computedValuesByCoord as Map<string, Map<number, unknown>>;
    expect(computedValuesByCoord.get("Sheet1")?.get(0)).toBe(123);
    expect(computedValuesByCoord.has("Budget")).toBe(false);

    (app as any).invalidateComputedValues([{ sheet: "Budget", address: "A1" }]);
    expect(computedValuesByCoord.get("Sheet1")?.has(0)).toBe(false);

    app.destroy();
    root.remove();
  });
});
