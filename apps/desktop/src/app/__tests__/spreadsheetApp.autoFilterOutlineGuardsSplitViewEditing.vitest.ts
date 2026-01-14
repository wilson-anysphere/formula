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

describe("SpreadsheetApp AutoFilter outline helpers respect split-view editing mode", () => {
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

  afterEach(() => {
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("no-ops AutoFilter outline mutations while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const outline = (app as any).getOutlineForSheet(app.getCurrentSheetId()) as any;
    expect(outline).toBeTruthy();

    // The legacy demo sheet seeds an outline group starting at index 2 (1-based).
    const entryIndex = 2; // row 1 (0-based) => outline index 2 (1-based)
    expect(outline.rows.entries.get(entryIndex)?.hidden?.filter).toBe(false);

    // Baseline: can apply filter-hidden state when not editing.
    app.setRowsFilteredHidden([1], true);
    expect(outline.rows.entries.get(entryIndex)?.hidden?.filter).toBe(true);

    // Split-view secondary editor owns edit mode via the desktop-shell global flag.
    (globalThis as any).__formulaSpreadsheetIsEditing = true;

    // All AutoFilter outline helpers should no-op while editing so they can't steal focus/mutate state.
    app.clearFilteredHiddenRowsInRange(0, 10);
    expect(outline.rows.entries.get(entryIndex)?.hidden?.filter).toBe(true);

    app.clearAllFilteredHiddenRows();
    expect(outline.rows.entries.get(entryIndex)?.hidden?.filter).toBe(true);

    // And new filtered-hidden state should not be applied during editing.
    const otherIndex = 3; // row 2 => outline index 3 (seeded by the demo outline group)
    expect(outline.rows.entries.get(otherIndex)?.hidden?.filter).toBe(false);
    app.setRowsFilteredHidden([2], true);
    expect(outline.rows.entries.get(otherIndex)?.hidden?.filter).toBe(false);

    app.destroy();
    root.remove();
  });
});

