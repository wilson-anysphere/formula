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

describe("SpreadsheetApp formula bar commit navigation", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Ensure tests default to legacy mode (simpler DOM/canvas mocking).
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
  });

  it("moves selection after Enter/Tab/Shift+Tab commits from the formula bar", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();

    // Tab moves right.
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    input!.focus();
    input!.value = "1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });

    // Shift+Tab moves left.
    input!.focus();
    input!.value = "2";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

    // Enter moves down.
    input!.focus();
    input!.value = "3";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 1, col: 0 });

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("commits and navigates on Tab/Shift+Tab/Enter even when the grid has focus (range selection mode)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();

    // Tab (grid focused) moves right.
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    input!.focus();
    input!.value = "1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    root.focus();
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });

    // Shift+Tab (grid focused) moves left.
    input!.focus();
    input!.value = "2";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    root.focus();
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });

    // Enter (grid focused) moves down.
    input!.focus();
    input!.value = "3";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    root.focus();
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 1, col: 0 });

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("navigates relative to the original edit cell when selection moved during range selection mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();

    // Begin editing A1.
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    input!.focus();
    input!.value = "1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));

    // Simulate the user moving the grid selection while still editing (e.g. picking a range).
    app.activateCell({ row: 4, col: 4 }, { scrollIntoView: false, focus: false });
    expect(app.getActiveCell()).toEqual({ row: 4, col: 4 });

    // Commit via Tab while the grid has focus. Navigation should be relative to the original edit
    // cell (A1 -> B1), not the transient selection (E5 -> F5).
    root.focus();
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("restores the original sheet before navigating on Tab commits (range selection mode)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    const doc = (app as any).document as { setCellValue: (sheetId: string, cell: { row: number; col: number }, value: unknown) => void };

    // Ensure Sheet2 exists so sheet switching is valid.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 7);

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();

    // Begin editing Sheet1!A1.
    expect((app as any).sheetId).toBe("Sheet1");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 0 });
    input!.focus();
    input!.value = "1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));

    // Switch to Sheet2 while still editing (simulates cross-sheet range picking / navigation).
    app.activateCell({ sheetId: "Sheet2", row: 0, col: 0 }, { scrollIntoView: false, focus: false });
    expect((app as any).sheetId).toBe("Sheet2");

    // Commit via Tab while the grid is focused. Should apply + navigate relative to Sheet1!A1,
    // restoring the sheet first.
    root.focus();
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect((app as any).sheetId).toBe("Sheet1");
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
