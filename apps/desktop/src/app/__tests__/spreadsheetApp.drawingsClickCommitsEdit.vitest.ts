/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import type { DrawingObject } from "../../drawings/types";
import { pxToEmu } from "../../drawings/overlay";

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
    }
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

function dispatchPointerEvent(
  target: HTMLElement,
  type: "pointerdown" | "pointerup",
  coords: { x: number; y: number },
  opts: { pointerId?: number; pointerType?: string; button?: number } = {}
): void {
  const event = new Event(type, { bubbles: true, cancelable: true }) as any;
  Object.defineProperty(event, "clientX", { value: coords.x });
  Object.defineProperty(event, "clientY", { value: coords.y });
  Object.defineProperty(event, "offsetX", { value: coords.x });
  Object.defineProperty(event, "offsetY", { value: coords.y });
  Object.defineProperty(event, "pointerId", { value: opts.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { value: opts.pointerType ?? "mouse" });
  Object.defineProperty(event, "button", { value: opts.button ?? 0 });
  target.dispatchEvent(event);
}

describe("SpreadsheetApp drawings click behavior while editing", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Ensure tests default to legacy mode unless explicitly overridden.
    priorGridMode = process.env.DESKTOP_GRID_MODE;
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

  it("commits an in-cell edit and selects the drawing when pointerdown hits a drawing", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "rect" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(50) },
        },
        zOrder: 0,
      },
    ];
    app.setDrawingObjects(objects);

    // Begin editing the active cell (A1).
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));
    const editor = root.querySelector<HTMLTextAreaElement>("textarea.cell-editor");
    expect(editor).not.toBeNull();
    editor!.value = "Hello";

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const cellRect = app.getCellRectA1("A1");
    expect(cellRect).not.toBeNull();

    // Hit inside the drawing bounds (anchored at A1 with 0 offset).
    const hitX = cellRect!.x + 10;
    const hitY = cellRect!.y + 10;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { x: hitX, y: hitY }, { pointerId: 1, pointerType: "mouse", button: 0 });
    dispatchPointerEvent(selectionCanvas, "pointerup", { x: hitX, y: hitY }, { pointerId: 1, pointerType: "mouse", button: 0 });

    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    expect(doc.getCell(sheetId, "A1").value).toBe("Hello");
    expect(app.getSelectedDrawingId()).toBe(1);

    app.destroy();
    root.remove();
  });

  it("commits an in-cell edit and selects the drawing when pointerdown hits a drawing (shared grid)", () => {
    process.env.DESKTOP_GRID_MODE = "shared";

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "rect" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(50) },
        },
        zOrder: 0,
      },
    ];
    app.setDrawingObjects(objects);

    // Begin editing the active cell (A1).
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));
    const editor = root.querySelector<HTMLTextAreaElement>("textarea.cell-editor");
    expect(editor).not.toBeNull();
    editor!.value = "Hello";

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const cellRect = app.getCellRectA1("A1");
    expect(cellRect).not.toBeNull();

    const hitX = cellRect!.x + 10;
    const hitY = cellRect!.y + 10;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { x: hitX, y: hitY }, { pointerId: 1, pointerType: "mouse", button: 0 });
    dispatchPointerEvent(selectionCanvas, "pointerup", { x: hitX, y: hitY }, { pointerId: 1, pointerType: "mouse", button: 0 });

    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    expect(doc.getCell(sheetId, "A1").value).toBe("Hello");
    expect(app.getSelectedDrawingId()).toBe(1);

    app.destroy();
    root.remove();
  });

  it("does not intercept drawing hits while the formula bar is in formula range selection mode", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const objects: DrawingObject[] = [
      {
        id: 7,
        kind: { type: "shape", label: "rect" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(50) },
        },
        zOrder: 0,
      },
    ];
    app.setDrawingObjects(objects);

    // Put the formula bar into formula editing mode (draft starts with "=").
    app.setFormulaBarDraft("=A1");
    expect(app.isFormulaBarFormulaEditing()).toBe(true);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const cellRect = app.getCellRectA1("A1");
    expect(cellRect).not.toBeNull();
    const hitX = cellRect!.x + 10;
    const hitY = cellRect!.y + 10;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { x: hitX, y: hitY }, { pointerId: 2, pointerType: "mouse", button: 0 });
    dispatchPointerEvent(selectionCanvas, "pointerup", { x: hitX, y: hitY }, { pointerId: 2, pointerType: "mouse", button: 0 });

    // Event should have bubbled to the grid root (drawing layer should not stop propagation),
    // and the drawing should not have been selected.
    expect(bubbled).toHaveBeenCalledTimes(1);
    expect(app.getSelectedDrawingId()).toBe(null);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("does not intercept drawing hits while the formula bar is in formula range selection mode (shared grid)", () => {
    process.env.DESKTOP_GRID_MODE = "shared";

    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });
    expect(app.getGridMode()).toBe("shared");

    const objects: DrawingObject[] = [
      {
        id: 7,
        kind: { type: "shape", label: "rect" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(50) },
        },
        zOrder: 0,
      },
    ];
    app.setDrawingObjects(objects);

    // Put the formula bar into formula editing mode (draft starts with "=").
    app.setFormulaBarDraft("=A1");
    expect(app.isFormulaBarFormulaEditing()).toBe(true);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const cellRect = app.getCellRectA1("A1");
    expect(cellRect).not.toBeNull();
    const hitX = cellRect!.x + 10;
    const hitY = cellRect!.y + 10;
    dispatchPointerEvent(selectionCanvas, "pointerdown", { x: hitX, y: hitY }, { pointerId: 2, pointerType: "mouse", button: 0 });
    dispatchPointerEvent(selectionCanvas, "pointerup", { x: hitX, y: hitY }, { pointerId: 2, pointerType: "mouse", button: 0 });

    // The drawings layer should not select the drawing or stop propagation during formula range selection.
    expect(bubbled).toHaveBeenCalledTimes(1);
    expect(app.getSelectedDrawingId()).toBe(null);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
