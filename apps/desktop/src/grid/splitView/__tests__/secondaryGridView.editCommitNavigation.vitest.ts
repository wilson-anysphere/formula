/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../secondaryGridView";
import { DocumentController } from "../../../document/documentController.js";
import type { ImageStore } from "../../../drawings/types";

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

describe("SecondaryGridView edit commit navigation", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0,
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

  const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

  it("commits cell edits and advances selection on Enter (A1 -> A2)", () => {
    const container = document.createElement("div");
    container.tabIndex = 0;
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 800 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 600 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    // Start editing A1 with an initial printable key press.
    container.dispatchEvent(new KeyboardEvent("keydown", { key: "h", bubbles: true, cancelable: true }));

    const editor = container.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor']");
    expect(editor).not.toBeNull();
    expect(editor?.classList.contains("cell-editor--open")).toBe(true);

    // Simulate the user typing "hello".
    if (!editor) throw new Error("Expected editor to exist");
    editor.value = "hello";

    // Commit with Enter. This used to throw at runtime due to a missing
    // `advanceSelectionAfterEdit` implementation.
    expect(() => {
      editor.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    }).not.toThrow();

    // Document edit was applied (A1 in sheet coords -> row=0,col=0).
    const cell = doc.getCell(sheetId, { row: 0, col: 0 }) as { value?: unknown } | null;
    expect(cell?.value).toBe("hello");

    // Selection advanced to A2 in grid coords (row=2,col=1 with 1x1 headers).
    expect(gridView.grid.renderer.getSelection()).toEqual({ row: 2, col: 1 });

    gridView.destroy();
    container.remove();
  });

  it("commits in-progress edits when clicking another cell in the grid", () => {
    const container = document.createElement("div");
    container.tabIndex = 0;
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 800 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 600 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    // Start editing A1.
    container.dispatchEvent(new KeyboardEvent("keydown", { key: "h", bubbles: true, cancelable: true }));

    const editor = container.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor']");
    expect(editor).not.toBeNull();
    if (!editor) throw new Error("Expected editor to exist");
    editor.value = "hello";

    // Click B1 (grid coords include header row/col at 0).
    const selectionCanvas = container.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();
    if (!selectionCanvas) throw new Error("Expected selection canvas to exist");

    selectionCanvas.dispatchEvent(
      new MouseEvent("pointerdown", {
        bubbles: true,
        cancelable: true,
        clientX: 48 + 100 + 10,
        clientY: 24 + 10,
        button: 0,
      }),
    );

    // Edit should have committed into A1 (sheet coords row=0,col=0).
    expect((doc.getCell(sheetId, { row: 0, col: 0 }) as any)?.value).toBe("hello");

    // Selection should have moved to B1 (grid coords row=1,col=2).
    expect(gridView.grid.renderer.getSelection()).toEqual({ row: 1, col: 2 });

    expect(editor.classList.contains("cell-editor--open")).toBe(false);

    gridView.destroy();
    container.remove();
  });

  it("commits edits on blur without stealing focus when focus moves outside the pane", () => {
    const container = document.createElement("div");
    container.tabIndex = 0;
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 800 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 600 });
    document.body.appendChild(container);

    const outside = document.createElement("button");
    outside.type = "button";
    outside.textContent = "outside";
    document.body.appendChild(outside);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 10,
      colCount: 10,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    // Start editing A1.
    container.dispatchEvent(new KeyboardEvent("keydown", { key: "h", bubbles: true, cancelable: true }));

    const editor = container.querySelector<HTMLTextAreaElement>("[data-testid='cell-editor']");
    expect(editor).not.toBeNull();
    if (!editor) throw new Error("Expected editor to exist");
    editor.value = "hello";
    expect(document.activeElement).toBe(editor);

    // Move focus elsewhere (e.g. clicking the ribbon, menus, or another panel). The blur-to-commit
    // handler should commit the edit but should *not* refocus the secondary pane (otherwise it would
    // steal focus from the clicked surface).
    outside.focus();

    expect((doc.getCell(sheetId, { row: 0, col: 0 }) as any)?.value).toBe("hello");
    expect(editor.classList.contains("cell-editor--open")).toBe(false);
    expect(document.activeElement).toBe(outside);

    gridView.destroy();
    container.remove();
    outside.remove();
  });
});
