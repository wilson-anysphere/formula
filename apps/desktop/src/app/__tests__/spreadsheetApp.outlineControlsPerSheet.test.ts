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
        // Default all unknown properties to no-op functions so rendering code can execute.
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

describe("SpreadsheetApp outline controls", () => {
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

  it("renders outline controls for the active sheet only", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const sheet1 = app.getCurrentSheetId();
    expect(sheet1).toBeTruthy();

    // The demo sheet seeds an outline group for rows 2-4 with a summary row at 5.
    const toggleSheet1 = root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-row-5"]');
    expect(toggleSheet1).toBeTruthy();
    expect(toggleSheet1?.textContent).toBe("-");

    // And columns 2-4 with a summary col at 5.
    const colToggleSheet1 = root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-col-5"]');
    expect(colToggleSheet1).toBeTruthy();
    expect(colToggleSheet1?.textContent).toBe("-");

    // Ensure Sheet2 exists, but with no outline groups by default.
    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    // Switching to Sheet2 should remove the outline toggle (no groups on that sheet).
    app.activateSheet("Sheet2");
    expect(root.querySelector('[data-testid="outline-toggle-row-5"]')).toBeNull();
    expect(root.querySelector('[data-testid="outline-toggle-col-5"]')).toBeNull();

    // Add an outline group on Sheet2 and re-render; the toggle should appear again.
    const outline2 = (app as any).getOutlineForSheet("Sheet2") as any;
    outline2.groupRows(2, 4);
    outline2.groupCols(2, 4);
    app.refresh();

    const toggleSheet2 = root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-row-5"]');
    expect(toggleSheet2).toBeTruthy();
    expect(toggleSheet2?.textContent).toBe("-");
    const colToggleSheet2 = root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-col-5"]');
    expect(colToggleSheet2).toBeTruthy();
    expect(colToggleSheet2?.textContent).toBe("-");

    // Collapse Sheet2's groups via the UI toggles.
    toggleSheet2?.click();
    expect(outline2.rows.entry(5).collapsed).toBe(true);
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-row-5"]')?.textContent).toBe("+");
    colToggleSheet2?.click();
    expect(outline2.cols.entry(5).collapsed).toBe(true);
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-col-5"]')?.textContent).toBe("+");

    // Switching back to Sheet1 should render Sheet1's toggle (still expanded).
    app.activateSheet(sheet1);
    const outline1 = (app as any).getOutlineForSheet(sheet1) as any;
    expect(outline1.rows.entry(5).collapsed).toBe(false);
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-row-5"]')?.textContent).toBe("-");
    expect(outline1.cols.entry(5).collapsed).toBe(false);
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-col-5"]')?.textContent).toBe("-");

    // Switching again to Sheet2 should retain Sheet2's collapsed state.
    app.activateSheet("Sheet2");
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-row-5"]')?.textContent).toBe("+");
    expect(root.querySelector<HTMLButtonElement>('[data-testid="outline-toggle-col-5"]')?.textContent).toBe("+");

    app.destroy();
    root.remove();
  });
});
