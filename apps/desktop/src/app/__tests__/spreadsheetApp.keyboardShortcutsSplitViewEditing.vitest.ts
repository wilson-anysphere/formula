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

describe("SpreadsheetApp keyboard shortcuts respect split-view editing mode", () => {
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

  it("no-ops auditing shortcuts while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    expect((app as any).auditingMode).toBe("off");

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "[", code: "BracketLeft", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect((app as any).auditingMode).toBe("precedents");

    app.clearAuditing();
    expect((app as any).auditingMode).toBe("off");

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "[", code: "BracketLeft", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect((app as any).auditingMode).toBe("off");

    app.destroy();
    root.remove();
  });

  it("no-ops show formulas shortcut while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    expect(app.getShowFormulas()).toBe(false);

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "`", code: "Backquote", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(app.getShowFormulas()).toBe(true);

    app.setShowFormulas(false);
    expect(app.getShowFormulas()).toBe(false);

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "`", code: "Backquote", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(app.getShowFormulas()).toBe(false);

    app.destroy();
    root.remove();
  });

  it("does not run workbook undo/redo while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Spy on the low-level undo/redo executor so the test remains independent of
    // DocumentController history internals.
    const applySpy = vi.spyOn(app as any, "applyUndoRedo").mockReturnValue(true);

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "z", code: "KeyZ", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(applySpy).toHaveBeenCalledTimes(1);

    applySpy.mockClear();
    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "z", code: "KeyZ", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(applySpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("no-ops formatting shortcuts while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc: any = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    const cell = { row: 0, col: 0 };

    // Force a stable single-cell selection.
    (app as any).selection = {
      type: "range",
      ranges: [{ startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col }],
      active: { ...cell },
      anchor: { ...cell },
      activeRangeIndex: 0,
    };

    expect(Boolean(doc.getCellFormat(sheetId, cell)?.font?.bold)).toBe(false);

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "b", code: "KeyB", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    expect(Boolean(doc.getCellFormat(sheetId, cell)?.font?.bold)).toBe(true);

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "b", code: "KeyB", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    // Without the split-view guard, this would toggle bold back off.
    expect(Boolean(doc.getCellFormat(sheetId, cell)?.font?.bold)).toBe(true);

    app.destroy();
    root.remove();
  });

  it("does not delete selected charts while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const deleteSpy = vi.spyOn((app as any).chartStore, "deleteChart");
    (app as any).selectedChartId = "chart-1";

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Delete", code: "Delete", bubbles: true, cancelable: true }),
    );
    expect(deleteSpy).toHaveBeenCalledTimes(1);

    deleteSpy.mockClear();
    (app as any).selectedChartId = "chart-2";
    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Delete", code: "Delete", bubbles: true, cancelable: true }),
    );
    expect(deleteSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("does not run clipboard shortcut handlers while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const copySpy = vi.spyOn(app as any, "copySelectionToClipboard").mockResolvedValue(undefined);
    const cutSpy = vi.spyOn(app as any, "cutSelectionToClipboard").mockResolvedValue(undefined);
    const pasteSpy = vi.spyOn(app as any, "pasteClipboardToSelection").mockResolvedValue(undefined);

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "c", code: "KeyC", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "x", code: "KeyX", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "v", code: "KeyV", ctrlKey: true, bubbles: true, cancelable: true }),
    );

    expect(copySpy).toHaveBeenCalledTimes(1);
    expect(cutSpy).toHaveBeenCalledTimes(1);
    expect(pasteSpy).toHaveBeenCalledTimes(1);

    copySpy.mockClear();
    cutSpy.mockClear();
    pasteSpy.mockClear();

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "c", code: "KeyC", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "x", code: "KeyX", ctrlKey: true, bubbles: true, cancelable: true }),
    );
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "v", code: "KeyV", ctrlKey: true, bubbles: true, cancelable: true }),
    );

    expect(copySpy).not.toHaveBeenCalled();
    expect(cutSpy).not.toHaveBeenCalled();
    expect(pasteSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("blocks F2 and Shift+F2 while split-view editing is active", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Avoid relying on grid geometry in this unit test: stub cell bounds and spy on editor open.
    vi.spyOn(app as any, "getCellRect").mockReturnValue({ x: 0, y: 0, width: 100, height: 20 });
    const editorOpenSpy = vi.spyOn((app as any).editor, "open").mockImplementation(() => {});

    expect(app.isCommentsPanelVisible()).toBe(false);

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2", code: "F2", bubbles: true, cancelable: true }));
    expect(editorOpenSpy).toHaveBeenCalledTimes(1);

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "F2", code: "F2", shiftKey: true, bubbles: true, cancelable: true }),
    );
    expect(app.isCommentsPanelVisible()).toBe(true);

    app.closeCommentsPanel();
    expect(app.isCommentsPanelVisible()).toBe(false);
    editorOpenSpy.mockClear();

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2", code: "F2", bubbles: true, cancelable: true }));
    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "F2", code: "F2", shiftKey: true, bubbles: true, cancelable: true }),
    );

    expect(editorOpenSpy).not.toHaveBeenCalled();
    expect(app.isCommentsPanelVisible()).toBe(false);

    app.destroy();
    root.remove();
  });
});
