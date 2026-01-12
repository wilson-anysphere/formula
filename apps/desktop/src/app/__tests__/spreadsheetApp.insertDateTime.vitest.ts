/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { dateToExcelSerial } from "../../shared/valueParsing.js";
import { SpreadsheetApp } from "../spreadsheetApp";

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

describe("SpreadsheetApp insert date/time shortcuts (Excel serial + numberFormat)", () => {
  afterEach(() => {
    vi.useRealTimers();
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

    vi.useFakeTimers();
    vi.setSystemTime(new Date(2020, 0, 2, 3, 4, 5));
  });

  it("Insert Date writes an Excel serial number and sets yyyy-mm-dd format for the selection", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    app.selectRange({ range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } }, { scrollIntoView: false, focus: true });

    const undoBefore = doc.getStackDepths().undo;
    app.insertDate();

    const expectedSerial = dateToExcelSerial(new Date(Date.UTC(2020, 0, 2)));
    expect(doc.getCell(sheetId, "A1").value).toBe(expectedSerial);
    expect(doc.getCellFormat(sheetId, "A1").numberFormat).toBe("yyyy-mm-dd");
    expect(doc.getStackDepths().undo).toBe(undoBefore + 1);
    expect(doc.undoLabel).toBe("Insert Date");

    const display = await app.getCellValueA1("A1");
    expect(display).toBe("2020-01-02");

    app.destroy();
    root.remove();
  });

  it("Insert Time writes a fractional-day serial number and sets hh:mm:ss format for the selection", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();

    app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } }, { scrollIntoView: false, focus: true });

    const undoBefore = doc.getStackDepths().undo;
    app.insertTime();

    const expectedSerial = (3 * 3600 + 4 * 60 + 5) / 86_400;
    expect(doc.getCell(sheetId, "A1").value).toBeCloseTo(expectedSerial, 10);
    expect(doc.getCellFormat(sheetId, "A1").numberFormat).toBe("hh:mm:ss");
    expect(doc.getStackDepths().undo).toBe(undoBefore + 1);
    expect(doc.undoLabel).toBe("Insert Time");

    const display = await app.getCellValueA1("A1");
    expect(display).toBe("03:04:05");

    app.destroy();
    root.remove();
  });
});

