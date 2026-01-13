/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../../app/spreadsheetApp";
import { computeFilterHiddenRows } from "../ribbonAutoFilter";

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

describe("Ribbon AutoFilter integration (legacy grid)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
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
  });

  it("applies filter state by hiding non-matching rows in the outline model", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    // A1:A4, with a header in A1.
    doc.setCellValue(sheetId, { row: 0, col: 0 }, "Header");
    doc.setCellValue(sheetId, { row: 1, col: 0 }, "x");
    doc.setCellValue(sheetId, { row: 2, col: 0 }, "y");
    doc.setCellValue(sheetId, { row: 3, col: 0 }, "x");

    const getValue = (row: number, col: number) => {
      const cell = doc.getCell(sheetId, { row, col }) as { value: unknown } | null;
      const value = cell?.value ?? null;
      return value == null ? "" : String(value);
    };

    const hidden = computeFilterHiddenRows({
      range: { startRow: 0, endRow: 3, startCol: 0, endCol: 0 },
      headerRows: 1,
      filterColumns: [{ colId: 0, values: ["x"] }],
      getValue,
    });

    // Clear any prior filter hidden flags and apply.
    app.clearFilteredHiddenRowsInRange(1, 3);
    app.setRowsFilteredHidden(hidden, true);

    const outline = (app as any).getOutlineForSheet(sheetId) as any;
    // Outline axis indices are 1-based (so row=2 => index=3).
    expect(outline.rows.entry(3).hidden.filter).toBe(true);
    expect(outline.rows.entry(2).hidden.filter).toBe(false);
    expect(outline.rows.entry(4).hidden.filter).toBe(false);

    app.destroy();
    root.remove();
  });
});
