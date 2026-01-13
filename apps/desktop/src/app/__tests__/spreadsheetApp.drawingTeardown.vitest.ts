/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../../drawings/imageBitmapCache";
import { DrawingOverlay, pxToEmu } from "../../drawings/overlay";
import { FormulaBarView } from "../../formula-bar/FormulaBarView";
import { AiCellFunctionEngine } from "../../spreadsheet/AiCellFunctionEngine.js";
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
  // JSDOM doesn't always implement pointer capture APIs.
  (root as any).setPointerCapture ??= () => {};
  (root as any).releasePointerCapture ??= () => {};
  document.body.appendChild(root);
  return root;
}

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
  const event =
    typeof (globalThis as any).PointerEvent === "function"
      ? new (globalThis as any).PointerEvent(type, base)
      : (() => {
          const e = new MouseEvent(type, base);
          Object.assign(e, { pointerId });
          return e;
        })();
  target.dispatchEvent(event);
}

describe("SpreadsheetApp drawings teardown", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

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
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 1 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("disposes drawing interaction listeners + clears bitmap caches on app.dispose()", () => {
    const root = createRoot();
    const formulaBarRoot = document.createElement("div");
    document.body.appendChild(formulaBarRoot);
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const clearSpy = vi.spyOn(ImageBitmapCache.prototype, "clear");
    const selectSpy = vi.spyOn(DrawingOverlay.prototype, "setSelectedId");
    const formulaBarDestroySpy = vi.spyOn(FormulaBarView.prototype, "destroy");
    const aiDisposeSpy = vi.spyOn(AiCellFunctionEngine.prototype, "dispose");
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true, formulaBar: formulaBarRoot });

    const outlineLayer = root.querySelector<HTMLDivElement>("div.outline-layer");
    const commentsThreads = root.querySelector<HTMLDivElement>("div.comments-panel__threads");
    const commentTooltip = root.querySelector<HTMLDivElement>("div.comment-tooltip");
    const auditingLegend = root.querySelector<HTMLDivElement>("div.auditing-legend");
    expect(outlineLayer).toBeTruthy();
    expect(commentsThreads).toBeTruthy();
    expect(commentTooltip).toBeTruthy();
    expect(auditingLegend).toBeTruthy();
    // Seed some children so dispose() can prove it clears detached DOM subtrees even when
    // the SpreadsheetApp instance stays referenced.
    outlineLayer!.appendChild(document.createElement("button"));
    commentsThreads!.appendChild(document.createElement("div"));
    commentTooltip!.appendChild(document.createElement("div"));
    auditingLegend!.appendChild(document.createElement("div"));

    // Register a few listeners and seed caches so destroy() can be validated as a
    // best-effort memory release path even when the app object stays referenced.
    app.subscribeSelection(() => {});
    app.subscribeScroll(() => {});
    app.subscribeZoom(() => {});
    app.onFormulaBarOverlayChange(() => {});
    app.onEditStateChange(() => {});
    app.subscribeDrawings(() => {});
    app.subscribeDrawingSelection(() => {});
    const imageStore = (app as any).imageStore as { set: (id: string, entry: any) => void; get: (id: string) => unknown };
    imageStore.set("in_cell_img", { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
    expect(imageStore.get("in_cell_img")).toBeTruthy();
    (app as any).computedValuesByCoord.set(app.getCurrentSheetId(), new Map([[0, 123]]));
    expect(((app as any).computedValuesByCoord as Map<string, unknown>).size).toBeGreaterThan(0);

    let interactionTarget =
      ((app as any).drawingInteractionController as { element?: EventTarget } | null)?.element ?? root;

    // Ensure the insert-image input (if created) is cleaned up on dispose.
    const input = (app as any).ensureInsertImageInput?.() as HTMLInputElement | undefined;
    expect(input).toBeTruthy();
    if (input) {
      input.onchange = () => {};
      expect(input.isConnected).toBe(true);
      expect((app as any).insertImageInput).toBe(input);
    }

    // Seed a single drawing object so pointer interactions have something to hit.
    const doc = app.getDocument() as any;
    doc.setSheetDrawings(app.getCurrentSheetId(), [
      {
        id: 1,
        zOrder: 0,
        kind: { type: "shape", label: "Box" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
          size: { cx: pxToEmu(120), cy: pxToEmu(80) },
        },
      },
    ]);

    const callbacks = (app as any).drawingInteractionCallbacks;
    expect(callbacks).toBeTruthy();
    const setObjectsSpy = vi.spyOn(callbacks, "setObjects");
    selectSpy.mockClear();

    // In shared-grid mode the DrawingInteractionController listens on the selection canvas
    // (the element that receives pointer events in the real UI).
    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).toBeTruthy();
    interactionTarget = selectionCanvas!;
    // JSDOM returns a default zero-sized rect for canvases; align with our stubbed root rect so
    // pointer coordinate conversion + hit-testing behaves consistently.
    (interactionTarget as any).getBoundingClientRect = root.getBoundingClientRect;
    (interactionTarget as any).setPointerCapture ??= () => {};
    (interactionTarget as any).releasePointerCapture ??= () => {};

    // Drag the object slightly: should call `setObjects`.
    dispatchPointerEvent(interactionTarget, "pointerdown", { clientX: 60, clientY: 40, pointerId: 1, button: 0 });
    // The pointerdown should also select the drawing (and not be immediately cleared by a redraw).
    expect(selectSpy.mock.calls.at(-1)?.[0]).toBe(1);
    dispatchPointerEvent(interactionTarget, "pointermove", { clientX: 80, clientY: 55, pointerId: 1 });
    expect(setObjectsSpy).toHaveBeenCalled();
    // End the gesture so `dispose()` doesn't need to cancel an in-flight drag (which can
    // legitimately call `setObjects` while restoring pre-gesture state).
    dispatchPointerEvent(interactionTarget, "pointerup", { clientX: 80, clientY: 55, pointerId: 1 });

    clearSpy.mockClear();

    app.dispose();

    // Overlay + caches should be cleared.
    expect(clearSpy).toHaveBeenCalled();
    expect(formulaBarDestroySpy).toHaveBeenCalled();
    expect(aiDisposeSpy).toHaveBeenCalled();
    expect(((app as any).computedValuesByCoord as Map<string, unknown>).size).toBe(0);
    expect(imageStore.get("in_cell_img")).toBeNull();
    expect(((app as any).selectionListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).scrollListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).zoomListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).formulaBarOverlayListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).editStateListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).drawingsListeners as Set<unknown>).size).toBe(0);
    expect(((app as any).drawingSelectionListeners as Set<unknown>).size).toBe(0);
    expect((app as any).dlpContext).toBeNull();
    expect(outlineLayer!.childElementCount).toBe(0);
    expect(commentsThreads!.childElementCount).toBe(0);
    expect(commentTooltip!.childElementCount).toBe(0);
    expect(auditingLegend!.childElementCount).toBe(0);
    if (input) {
      expect(input.isConnected).toBe(false);
      expect(input.onchange).toBeNull();
      expect((app as any).insertImageInput).toBeNull();
    }

    // Disposing may cancel an in-flight gesture by restoring the pre-drag object list. Clear spies
    // after dispose so we only assert on *post-dispose* pointer events.
    setObjectsSpy.mockClear();

    dispatchPointerEvent(interactionTarget, "pointerdown", { clientX: 60, clientY: 40, pointerId: 2, button: 0 });
    dispatchPointerEvent(interactionTarget, "pointermove", { clientX: 100, clientY: 70, pointerId: 2 });
    expect(setObjectsSpy).not.toHaveBeenCalled();

    root.remove();
  });
});
