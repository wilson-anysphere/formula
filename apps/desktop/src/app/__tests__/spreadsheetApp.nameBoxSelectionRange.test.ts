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

describe("SpreadsheetApp formula bar name box selection range", () => {
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

  it("shows the selection range (A1:B3) in the formula bar name box", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    app.selectRange(
      {
        range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }, // A1:B3
      },
      { scrollIntoView: false, focus: false },
    );

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("A1:B3");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("updates the name box while the formula bar is editing", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const formulaInput = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(formulaInput).not.toBeNull();
    formulaInput!.focus();

    app.selectRange(
      {
        range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }, // A1:B3
      },
      { scrollIntoView: false, focus: false },
    );

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("A1:B3");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("does not overwrite the name box while the user is typing", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();

    address!.focus();
    address!.value = "Z99";

    // Selection changes should not clobber the user's in-progress typing.
    app.selectRange(
      {
        range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 }, // A1:B3
      },
      { scrollIntoView: false, focus: false },
    );
    expect(address!.value).toBe("Z99");

    // Esc should revert to the latest selection display value.
    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(address!.value).toBe("A1:B3");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("shows a stable label for multi-range selections", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    // Legacy grid mode doesn't expose a public multi-range selection API, but the
    // underlying selection model supports it (and status/formula bar should render it).
    (app as any).selection = buildSelection(
      {
        ranges: [
          { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, // A1
          { startRow: 2, endRow: 2, startCol: 2, endCol: 2 }, // C3
        ],
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        activeRangeIndex: 0,
      },
      (app as any).limits,
    );
    (app as any).updateStatus();

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("2 ranges");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("formats full-column selection like Excel (A:A)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    const limits = (app as any).limits as { maxRows: number; maxCols: number };

    app.selectRange(
      {
        range: { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: 0 }, // Column A
      },
      { scrollIntoView: false, focus: false },
    );

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("A:A");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("formats full-row selection like Excel (1:1)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    const limits = (app as any).limits as { maxRows: number; maxCols: number };

    app.selectRange(
      {
        range: { startRow: 0, endRow: 0, startCol: 0, endCol: limits.maxCols - 1 }, // Row 1
      },
      { scrollIntoView: false, focus: false },
    );

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("1:1");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("treats the multi-range name box label (N ranges) as a no-op go to", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    // Create a multi-range selection and force a status update so the name box shows "2 ranges".
    (app as any).selection = buildSelection(
      {
        ranges: [
          { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, // A1
          { startRow: 2, endRow: 2, startCol: 2, endCol: 2 }, // C3
        ],
        active: { row: 0, col: 0 },
        anchor: { row: 0, col: 0 },
        activeRangeIndex: 0,
      },
      (app as any).limits,
    );
    (app as any).updateStatus();

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();
    expect(address!.value).toBe("2 ranges");

    address!.focus();
    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    // Should not trigger invalid-reference feedback for the display label.
    expect(address!.getAttribute("aria-invalid")).not.toBe("true");
    expect(status.selectionRange.textContent).toBe("2 ranges");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("navigates to full-column and full-row references entered into the name box", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const address = formulaBar.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();

    // Go to column B (B:B).
    address!.focus();
    address!.value = "B:B";
    address!.dispatchEvent(new Event("input", { bubbles: true }));
    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    expect(status.selectionRange.textContent).toBe("B:B");

    // Go to row 2 (2:2).
    address!.focus();
    address!.value = "2:2";
    address!.dispatchEvent(new Event("input", { bubbles: true }));
    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    expect(status.selectionRange.textContent).toBe("2:2");

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
