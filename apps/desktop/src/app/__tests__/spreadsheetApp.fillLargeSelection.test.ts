/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

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

describe("SpreadsheetApp fill large selections", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

    document.body.innerHTML = "";

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
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("blocks fill down when the target area is extremely large", () => {
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

    // 5000x100 = 500,000 target cells for fill down (well above MAX_FILL_CELLS).
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: 0, endRow: 4_999, startCol: 0, endCol: 99 }],
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      activeRangeIndex: 0,
    };

    app.fillDown();
    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("blocks shared-grid fill handle commits when the delta range is extremely large", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const doc = app.getDocument();
    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    // Grid ranges include a 1-row/1-col header at index 0.
    // Source: A1 (1 cell). Target delta: A2:A200002 (200,001 cells) => exceeds MAX_FILL_CELLS.
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 200_003, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Allow any microtasks scheduled by the guard path to flush.
    await Promise.resolve();

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("blocks shared-grid fill handle commits while the formula bar is actively editing", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    expect(app.getGridMode()).toBe("shared");

    const doc = app.getDocument();
    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();
    input!.focus();
    expect(app.isEditing()).toBe(true);

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    // Grid ranges include a 1-row/1-col header at index 0.
    // Source: A1 (1 cell). Target delta: A2 (1 cell).
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Allow any microtasks scheduled by the guard path to flush.
    await Promise.resolve();
    await Promise.resolve();

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("shows an encryption-aware toast and restores selection when shared-grid fill handle is blocked by canEditCell", async () => {
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const doc = app.getDocument();
    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    // Simulate an encryption/permission guard installed by collab mode.
    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    (app as any).collabSession = {
      cells,
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    const initialSelection = [{ startRow: 1, endRow: 2, startCol: 1, endCol: 2 }];
    sharedGrid.setSelectionRanges(initialSelection, { activeIndex: 0, activeCell: { row: 1, col: 1 }, scrollIntoView: false });

    // Grid ranges include a 1-row/1-col header at index 0.
    // Source: A1 (1 cell). Target delta: A2 (1 cell).
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
      mode: "formulas",
    });

    // Simulate DesktopSharedGrid expanding selection after the callback returns.
    sharedGrid.setSelectionRanges(
      [{ startRow: 1, endRow: 3, startCol: 1, endCol: 2 }],
      { activeIndex: 0, activeCell: { row: 2, col: 1 }, scrollIntoView: false },
    );

    // Allow any microtasks scheduled by the guard path to flush.
    await Promise.resolve();
    await Promise.resolve();

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    expect(sharedGrid.renderer.getSelectionRanges()).toEqual(initialSelection);

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Missing encryption key");

    app.destroy();
    root.remove();
  });

  it("blocks shared-grid fill handle commits when the shell reports split-view editing", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");
    expect(app.isEditing()).toBe(false);

    const doc = app.getDocument();
    const beginBatch = vi.spyOn(doc, "beginBatch");
    const setCellInput = vi.spyOn(doc, "setCellInput");

    (globalThis as any).__formulaSpreadsheetIsEditing = true;

    const sharedGrid = (app as any).sharedGrid as any;
    const onFillCommit = sharedGrid?.callbacks?.onFillCommit as ((event: any) => void) | undefined;
    expect(typeof onFillCommit).toBe("function");

    // Grid ranges include a 1-row/1-col header at index 0.
    // Source: A1 (1 cell). Target delta: A2 (1 cell).
    onFillCommit!({
      sourceRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
      targetRange: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
      mode: "copy",
    });

    // Allow any microtasks scheduled by the guard path to flush.
    await Promise.resolve();
    await Promise.resolve();

    expect(beginBatch).not.toHaveBeenCalled();
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("blocks legacy fill-drag commits when the target area is extremely large", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const doc = app.getDocument();
    const setCellInput = vi.spyOn(doc, "setCellInput");

    const applied = (app as any).applyFill(
      { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
      { startRow: 0, endRow: 300_000, startCol: 0, endCol: 0 },
      "formulas",
    );

    expect(applied).toBe(false);
    expect(setCellInput).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
