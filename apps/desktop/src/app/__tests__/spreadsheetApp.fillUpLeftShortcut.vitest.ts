/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

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

describe("SpreadsheetApp fill up/left shortcuts", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  it("fills up by copying the bottom row into rows above", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setRangeValues(
      sheetId,
      "A1",
      [
        ["a1", "b1"],
        ["a2", "b2"],
        ["a3", "b3"],
      ],
      { label: "Seed" },
    );

    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 2, startCol: 0, endCol: 1 }], // A1:B3
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillUp();
    await app.whenIdle();

    const a1 = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
    const b1 = doc.getCell(sheetId, { row: 0, col: 1 }) as any;
    const a2 = doc.getCell(sheetId, { row: 1, col: 0 }) as any;
    const b2 = doc.getCell(sheetId, { row: 1, col: 1 }) as any;
    const a3 = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
    const b3 = doc.getCell(sheetId, { row: 2, col: 1 }) as any;

    expect(a1.value).toBe("a3");
    expect(b1.value).toBe("b3");
    expect(a2.value).toBe("a3");
    expect(b2.value).toBe("b3");
    expect(a3.value).toBe("a3");
    expect(b3.value).toBe("b3");

    app.destroy();
    root.remove();
  });

  it("fills left by copying the rightmost column into columns to the left", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setRangeValues(
      sheetId,
      "A1",
      [
        ["a1", "b1", "c1"],
        ["a2", "b2", "c2"],
      ],
      { label: "Seed" },
    );

    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }], // A1:C2
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillLeft();
    await app.whenIdle();

    const a1 = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
    const b1 = doc.getCell(sheetId, { row: 0, col: 1 }) as any;
    const c1 = doc.getCell(sheetId, { row: 0, col: 2 }) as any;
    const a2 = doc.getCell(sheetId, { row: 1, col: 0 }) as any;
    const b2 = doc.getCell(sheetId, { row: 1, col: 1 }) as any;
    const c2 = doc.getCell(sheetId, { row: 1, col: 2 }) as any;

    expect(a1.value).toBe("c1");
    expect(b1.value).toBe("c1");
    expect(c1.value).toBe("c1");
    expect(a2.value).toBe("c2");
    expect(b2.value).toBe("c2");
    expect(c2.value).toBe("c2");

    app.destroy();
    root.remove();
  });

  it("rewrites formulas with negative deltas when filling up/left (engine rewrite hook)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    // Fill Up: A2 -> A1 (deltaRow: -1).
    doc.setRangeValues(sheetId, "A1", [[null], [{ formula: "=SUM(2:2)" }]], { label: "Seed" });
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }], // A1:A2
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    const rewrite = vi.fn(async (requests: Array<{ formula: string; deltaRow: number; deltaCol: number }>) => {
      return requests.map((r) => {
        if (r.formula === "=SUM(2:2)" && r.deltaRow === -1 && r.deltaCol === 0) {
          return "=SUM(1:1)";
        }
        if (r.formula === "=SUM(C:C)" && r.deltaRow === 0 && r.deltaCol === -1) {
          return "=SUM(B:B)";
        }
        if (r.formula === "=SUM(C:C)" && r.deltaRow === 0 && r.deltaCol === -2) {
          return "=SUM(A:A)";
        }
        return r.formula;
      });
    });

    (app as any).wasmEngine = { rewriteFormulasForCopyDelta: rewrite, terminate: () => {} };

    app.fillUp();
    await app.whenIdle();

    expect(rewrite).toHaveBeenCalled();
    expect(rewrite.mock.calls[0]?.[0]).toEqual([{ formula: "=SUM(2:2)", deltaRow: -1, deltaCol: 0 }]);

    const a1 = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
    expect(a1.formula).toBe("=SUM(1:1)");

    // Fill Left: C1 -> A1:B1 (deltaCol: -2 / -1).
    doc.setRangeValues(sheetId, "A1", [[null, null, { formula: "=SUM(C:C)" }]], { label: "Seed2" });
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 2 }], // A1:C1
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillLeft();
    await app.whenIdle();

    // Second invocation: fillLeft.
    expect(rewrite.mock.calls[1]?.[0]).toEqual([
      { formula: "=SUM(C:C)", deltaRow: 0, deltaCol: -2 },
      { formula: "=SUM(C:C)", deltaRow: 0, deltaCol: -1 },
    ]);

    const leftA1 = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
    const leftB1 = doc.getCell(sheetId, { row: 0, col: 1 }) as any;
    expect(leftA1.formula).toBe("=SUM(A:A)");
    expect(leftB1.formula).toBe("=SUM(B:B)");

    app.destroy();
    root.remove();
  });

  it("blocks fill up when the target area is extremely large", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    // 5000x100 = 500,000 target cells for fill up (well above MAX_FILL_CELLS).
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 4_999, startCol: 0, endCol: 99 }],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillUp();
    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("blocks fill up/left when the workbook is read-only (collab viewer)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setRangeValues(sheetId, "A1", [["seed"]], { label: "Seed" });

    // Minimal stub so `SpreadsheetApp.isReadOnly()` returns true.
    (app as any).collabSession = { isReadOnly: () => true };

    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }], // A1:A2
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillUp();
    app.fillLeft();

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    const a1 = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
    expect(a1.value).toBe("seed");

    app.destroy();
    root.remove();
  });
});
