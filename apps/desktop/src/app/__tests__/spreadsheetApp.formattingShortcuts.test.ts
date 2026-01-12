/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { buildSelection } from "../../selection/selection";

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

describe("SpreadsheetApp formatting keyboard shortcuts", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Needed for showToast(...) if it fires (e.g. safety cap).
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("Ctrl/Cmd+B toggles bold across all selection ranges", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    const rangeA1 = { startRow: 0, endRow: 0, startCol: 0, endCol: 0 };
    const rangeC3 = { startRow: 2, endRow: 2, startCol: 2, endCol: 2 };
    (app as any).selection = buildSelection(
      { ranges: [rangeA1, rangeC3], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, activeRangeIndex: 0 },
      (app as any).limits,
    );

    const event = new KeyboardEvent("keydown", { key: "b", ctrlKey: true, cancelable: true });
    root.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);

    for (const cellCoord of [
      { row: 0, col: 0 },
      { row: 2, col: 2 },
    ]) {
      const cell = doc.getCell(sheetId, cellCoord) as any;
      expect(cell.styleId).not.toBe(0);
      const style = doc.styleTable.get(cell.styleId) as any;
      expect(style.font?.bold).toBe(true);
    }

    app.destroy();
    root.remove();
  });

  it("Ctrl/Cmd+U toggles underline on the current selection range", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    app.selectRange({ range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } }); // A1:B2

    const event = new KeyboardEvent("keydown", { key: "u", ctrlKey: true, cancelable: true });
    root.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);

    for (let row = 0; row <= 1; row += 1) {
      for (let col = 0; col <= 1; col += 1) {
        const cell = doc.getCell(sheetId, { row, col }) as any;
        const style = doc.styleTable.get(cell.styleId) as any;
        expect(style.font?.underline).toBe(true);
      }
    }

    app.destroy();
    root.remove();
  });

  it("Ctrl+I toggles italic, but Cmd+I is not captured", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    // Ctrl+I should apply italic.
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } }); // A1
    const ctrlEvent = new KeyboardEvent("keydown", { key: "i", ctrlKey: true, cancelable: true });
    root.dispatchEvent(ctrlEvent);
    expect(ctrlEvent.defaultPrevented).toBe(true);
    {
      const cell = doc.getCell(sheetId, { row: 0, col: 0 }) as any;
      const style = doc.styleTable.get(cell.styleId) as any;
      expect(style.font?.italic).toBe(true);
    }

    // Cmd+I should *not* be captured (reserved for the AI sidebar).
    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 1, endCol: 1 } }); // B1
    const cmdEvent = new KeyboardEvent("keydown", { key: "i", metaKey: true, cancelable: true });
    root.dispatchEvent(cmdEvent);
    expect(cmdEvent.defaultPrevented).toBe(false);
    {
      const cell = doc.getCell(sheetId, { row: 0, col: 1 }) as any;
      expect(cell.styleId).toBe(0);
    }

    app.destroy();
    root.remove();
  });

  it.each([
    ["Ctrl/Cmd+Shift+$ applies currency format", { key: "$", code: "Digit4", preset: "$#,##0.00" }],
    ["Ctrl/Cmd+Shift+% applies percent format", { key: "%", code: "Digit5", preset: "0%" }],
    ["Ctrl/Cmd+Shift+# applies date format", { key: "#", code: "Digit3", preset: "m/d/yyyy" }],
  ])("%s", (_name, cfg) => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 1 } }); // A1:B1

    const event = new KeyboardEvent("keydown", {
      key: cfg.key,
      code: cfg.code,
      ctrlKey: true,
      shiftKey: true,
      cancelable: true,
    });
    root.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);

    for (let col = 0; col <= 1; col += 1) {
      const cell = doc.getCell(sheetId, { row: 0, col }) as any;
      const style = doc.styleTable.get(cell.styleId) as any;
      expect(style.numberFormat).toBe(cfg.preset);
    }

    app.destroy();
    root.remove();
  });
});

