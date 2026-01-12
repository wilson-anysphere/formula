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

describe("SpreadsheetApp edit state API", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Ensure tests default to legacy mode unless explicitly overridden.
    process.env.DESKTOP_GRID_MODE = "legacy";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
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

  it("fires edit state changes for the in-grid cell editor (F2 + Enter/Escape)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const states: boolean[] = [];
    const unsubscribe = app.onEditStateChange((isEditing) => states.push(isEditing));

    // Immediate fire.
    expect(states).toEqual([false]);
    expect(app.isEditing()).toBe(false);

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));
    expect(app.isEditing()).toBe(true);
    expect(states).toEqual([false, true]);

    const editor = root.querySelector("textarea.cell-editor");
    expect(editor).not.toBeNull();

    // Commit via Enter.
    (editor as HTMLTextAreaElement).value = "Hello";
    editor!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(app.isEditing()).toBe(false);
    expect(states).toEqual([false, true, false]);

    // Re-open and cancel via Escape.
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));
    expect(app.isEditing()).toBe(true);
    expect(states).toEqual([false, true, false, true]);

    const editor2 = root.querySelector("textarea.cell-editor");
    expect(editor2).not.toBeNull();
    editor2!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(app.isEditing()).toBe(false);
    expect(states).toEqual([false, true, false, true, false]);

    unsubscribe();
    app.destroy();
    root.remove();
  });

  it("opens the in-grid cell editor with the caret at the end of the existing value (F2 semantics)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    doc.setCellValue(sheetId, "A1", "hello");

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));

    const editor = root.querySelector<HTMLTextAreaElement>("textarea.cell-editor");
    expect(editor).not.toBeNull();
    expect(editor!.value).toBe("hello");

    const end = editor!.value.length;
    expect(editor!.selectionStart).toBe(end);
    expect(editor!.selectionEnd).toBe(end);

    editor!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(app.isEditing()).toBe(false);

    app.destroy();
    root.remove();
  });

  it("fires edit state changes for formula bar begin edit + commit/cancel", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const states: boolean[] = [];
    app.onEditStateChange((isEditing) => states.push(isEditing));

    expect(states).toEqual([false]);

    const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
    expect(input).not.toBeNull();

    // Begin edit.
    input!.focus();
    expect(app.isEditing()).toBe(true);
    expect(states).toEqual([false, true]);

    // Commit via Enter.
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(app.isEditing()).toBe(false);
    expect(states).toEqual([false, true, false]);

    // Begin edit again.
    input!.focus();
    expect(app.isEditing()).toBe(true);
    expect(states).toEqual([false, true, false, true]);

    // Cancel via Escape.
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(app.isEditing()).toBe(false);
    expect(states).toEqual([false, true, false, true, false]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
