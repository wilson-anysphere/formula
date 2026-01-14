/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { rangeToA1 } from "../../selection/a1";

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

describe("SpreadsheetApp formula bar keyboard point mode", () => {
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

  it("Arrow keys update formula references while grid has focus during formula editing", () => {
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

    // Start editing a formula in the formula bar.
    input!.focus();
    input!.value = "=SUM(";
    input!.setSelectionRange(input!.value.length, input!.value.length);
    input!.dispatchEvent(new Event("input", { bubbles: true }));

    // Simulate Excel-style range selection mode: focus returns to grid.
    root.focus();

    const right = new KeyboardEvent("keydown", { key: "ArrowRight", cancelable: true });
    root.dispatchEvent(right);
    expect(right.defaultPrevented).toBe(true);
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });
    expect(input!.value).toBe("=SUM(B1");

    const down = new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true });
    root.dispatchEvent(down);
    expect(down.defaultPrevented).toBe(true);
    expect(app.getActiveCell()).toEqual({ row: 1, col: 1 });
    expect(input!.value).toBe("=SUM(B2");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("Shift+Arrow extends range and inserts A1:B1 references while grid has focus", () => {
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

    input!.focus();
    input!.value = "=SUM(";
    input!.setSelectionRange(input!.value.length, input!.value.length);
    input!.dispatchEvent(new Event("input", { bubbles: true }));

    root.focus();

    const event = new KeyboardEvent("keydown", { key: "ArrowRight", shiftKey: true, cancelable: true });
    root.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    expect(app.getActiveCell()).toEqual({ row: 0, col: 1 });

    const expectedRange = rangeToA1({ startRow: 0, endRow: 0, startCol: 0, endCol: 1 });
    expect(input!.value).toBe(`=SUM(${expectedRange}`);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});

