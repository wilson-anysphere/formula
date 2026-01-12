/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

describe("SpreadsheetApp internal paste detection (RTF)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // SpreadsheetApp schedules renders via requestAnimationFrame.
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

  it("treats a paste as internal when clipboard content matches only via rtf (preserving styleId + shifting formulas)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();

    const styleId = doc.styleTable.intern({ font: { bold: true } });
    doc.setRangeValues(app.getCurrentSheetId(), "A1", [[{ formula: "=A2", styleId }]], { label: "Seed" });

    const copiedRange = { startRow: 0, endRow: 0, startCol: 0, endCol: 0 };
    const cells = (app as any).snapshotClipboardCells(copiedRange);

    const ctxRtf = "{\\rtf1\\ansi Hello}\r\n";
    const contentRtf = "{\\rtf1\\ansi Hello}\n\n";

    (app as any).clipboardCopyContext = {
      range: copiedRange,
      payload: { rtf: ctxRtf },
      cells,
    };

    const provider = {
      read: vi.fn(async () => ({ rtf: contentRtf })),
      write: vi.fn(async () => {}),
    };
    (app as any).clipboardProviderPromise = Promise.resolve(provider);

    app.activateCell({ row: 1, col: 1 }); // B2

    await (app as any).pasteClipboardToSelection();

    const pasted = doc.getCell(app.getCurrentSheetId(), { row: 1, col: 1 }) as any;
    expect(pasted.styleId).toBe(styleId);
    expect(pasted.formula).toBe("=B3");

    app.destroy();
    root.remove();
  });
});

