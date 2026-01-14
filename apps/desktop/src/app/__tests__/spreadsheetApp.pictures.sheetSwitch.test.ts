/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { pxToEmu } from "../../drawings/overlay";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

function dispatchPointer(target: EventTarget, type: string, point: { x: number; y: number; pointerId?: number }): void {
  const event = new Event(type, { bubbles: true, cancelable: true }) as any;
  const pointerId = point.pointerId ?? 1;
  Object.defineProperties(event, {
    clientX: { value: point.x, configurable: true },
    clientY: { value: point.y, configurable: true },
    offsetX: { value: point.x, configurable: true },
    offsetY: { value: point.y, configurable: true },
    pointerId: { value: pointerId, configurable: true },
    pointerType: { value: "mouse", configurable: true },
    button: { value: 0, configurable: true },
    buttons: { value: 1, configurable: true },
  });
  target.dispatchEvent(event);
}

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
      drawImage: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        // Default all unknown properties to no-op functions so rendering code can execute.
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

function deleteSeededCharts(app: SpreadsheetApp): void {
  // Canvas charts are enabled by default, so ChartStore charts appear in drawings APIs
  // (e.g. `getDrawingsDebugState`). Remove any charts so these tests can focus on picture counts
  // and sheet switching behavior without extra drawing-layer objects.
  for (const chart of app.listCharts()) {
    (app as any).chartStore.deleteChart(chart.id);
  }
}

describe("SpreadsheetApp pictures/drawings sheet switching", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    // Exercise shared-grid mode so the sheet switch path resets both drawing
    // selection and the shared-grid drawing interaction controller.
    process.env.DESKTOP_GRID_MODE = "shared";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("scopes pictures per sheet and clears drawing selection on sheet switch", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    deleteSeededCharts(app);
    const file = new File([new Uint8Array([1, 2, 3, 4])], "cat.png", { type: "image/png" });
    await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

    const sheet1Initial = app.getDrawingsDebugState();
    expect(sheet1Initial.sheetId).toBe("Sheet1");
    const sheet1ImagesInitial = sheet1Initial.drawings.filter((d) => d.kind === "image");
    expect(sheet1ImagesInitial).toHaveLength(1);
    const insertedId = sheet1ImagesInitial[0]!.id;
    expect(sheet1Initial.selectedId).toBe(insertedId);

    // Ensure Sheet2 exists.
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    app.activateSheet("Sheet2");
    const sheet2 = app.getDrawingsDebugState();
    expect(sheet2.sheetId).toBe("Sheet2");
    expect(sheet2.drawings).toHaveLength(0);
    expect(sheet2.selectedId).toBe(null);

    app.activateSheet("Sheet1");
    const sheet1After = app.getDrawingsDebugState();
    expect(sheet1After.sheetId).toBe("Sheet1");
    const sheet1ImagesAfter = sheet1After.drawings.filter((d) => d.kind === "image");
    expect(sheet1ImagesAfter).toHaveLength(1);
    expect(sheet1ImagesAfter[0]?.id).toBe(insertedId);
    // Selection should not "carry over" when switching back.
    expect(sheet1After.selectedId).toBe(null);

    app.destroy();
    root.remove();
  });

  it("cancels an in-progress picture drag when switching sheets (no leakage across sheets)", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Enable the shared-grid drawing interaction controller so pointer gestures
    // would normally commit via `commitObjects` on pointerup.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    deleteSeededCharts(app);

    const file = new File([new Uint8Array([1, 2, 3, 4])], "cat.png", { type: "image/png" });
    await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

    const sheet1Initial = app.getDrawingsDebugState();
    expect(sheet1Initial.sheetId).toBe("Sheet1");
    const sheet1ImagesInitial = sheet1Initial.drawings.filter((d) => d.kind === "image");
    expect(sheet1ImagesInitial).toHaveLength(1);
    const inserted = sheet1ImagesInitial[0]!;
    expect(inserted.rectPx).not.toBeNull();

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const start = {
      x: inserted.rectPx!.x + inserted.rectPx!.width / 2,
      y: inserted.rectPx!.y + inserted.rectPx!.height / 2,
    };
    const move = { x: start.x + 40, y: start.y + 20 };

    dispatchPointer(selectionCanvas, "pointerdown", { ...start, pointerId: 1 });
    // Verify the controller is actively dragging so the sheet switch path must cancel it.
    const controller = (app as any).drawingInteractionController as any;
    expect(controller?.dragging).not.toBeNull();

    dispatchPointer(selectionCanvas, "pointermove", { ...move, pointerId: 1 });

    // Ensure Sheet2 exists.
    app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    // Switch sheets while the pointer is still down (mid-gesture).
    app.activateSheet("Sheet2");
    expect((app as any).drawingInteractionController?.dragging ?? null).toBeNull();

    // Release the pointer after the sheet switch. This should not commit the picture into Sheet2.
    dispatchPointer(selectionCanvas, "pointerup", { ...move, pointerId: 1 });

    const sheet2 = app.getDrawingsDebugState();
    expect(sheet2.sheetId).toBe("Sheet2");
    expect(sheet2.drawings).toHaveLength(0);
    expect(sheet2.selectedId).toBe(null);

    const docAny = app.getDocument() as any;
    const rawSheet2 = (() => {
      try {
        return docAny.getSheetDrawings?.("Sheet2");
      } catch {
        return null;
      }
    })();
    expect(Array.isArray(rawSheet2) ? rawSheet2 : []).toHaveLength(0);

    app.activateSheet("Sheet1");
    const sheet1After = app.getDrawingsDebugState();
    expect(sheet1After.sheetId).toBe("Sheet1");
    expect(sheet1After.drawings.filter((d) => d.kind === "image")).toHaveLength(1);
    expect(sheet1After.selectedId).toBe(null);

    app.destroy();
    root.remove();
  });

  it("cancels an in-progress picture drag in legacy mode when switching sheets (no cross-sheet commit)", async () => {
    // Override the shared-grid default so this test exercises SpreadsheetApp's legacy drawing gesture state machine.
    process.env.DESKTOP_GRID_MODE = "legacy";

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    deleteSeededCharts(app);
    const doc: any = app.getDocument();

    const file = new File([new Uint8Array([1, 2, 3, 4])], "cat.png", { type: "image/png" });
    await app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

    const sheet1Initial = app.getDrawingsDebugState();
    expect(sheet1Initial.sheetId).toBe("Sheet1");
    const sheet1ImagesInitial = sheet1Initial.drawings.filter((d) => d.kind === "image");
    expect(sheet1ImagesInitial).toHaveLength(1);
    const inserted = sheet1ImagesInitial[0]!;
    expect(inserted.rectPx).not.toBeNull();

    // Ensure Sheet2 exists and contains a drawing with the *same* id. Without canceling the in-flight gesture,
    // the pointerup commit would update this drawing after the sheet id changes (cross-sheet leak).
    doc.setSheetDrawings("Sheet2", [
      {
        id: String(inserted.id),
        kind: { type: "shape", label: "Sheet2 Box" },
        anchor: { type: "absolute", pos: { xEmu: pxToEmu(10), yEmu: pxToEmu(10) }, size: { cx: pxToEmu(80), cy: pxToEmu(40) } },
        zOrder: 0,
      },
    ]);

    const sheet2Before = Array.isArray(doc.getSheetDrawings("Sheet2")) ? doc.getSheetDrawings("Sheet2") : [];
    expect(sheet2Before).toHaveLength(1);
    const sheet2BeforeAnchor = JSON.parse(JSON.stringify((sheet2Before[0] as any).anchor));

    const start = {
      x: inserted.rectPx!.x + inserted.rectPx!.width / 2,
      y: inserted.rectPx!.y + inserted.rectPx!.height / 2,
    };
    const move = { x: start.x + 40, y: start.y + 20 };

    dispatchPointer(root, "pointerdown", { ...start, pointerId: 42 });
    dispatchPointer(root, "pointermove", { ...move, pointerId: 42 });
    expect((app as any).drawingGesture).not.toBeNull();

    app.activateSheet("Sheet2");
    expect((app as any).drawingGesture ?? null).toBeNull();

    // Release pointer after the sheet switch. This must not update the drawing on Sheet2.
    dispatchPointer(root, "pointerup", { ...move, pointerId: 42 });

    const sheet2After = Array.isArray(doc.getSheetDrawings("Sheet2")) ? doc.getSheetDrawings("Sheet2") : [];
    expect(sheet2After).toHaveLength(1);
    expect((sheet2After[0] as any).anchor).toEqual(sheet2BeforeAnchor);

    // And the original picture should still be on Sheet1.
    app.activateSheet("Sheet1");
    const sheet1After = app.getDrawingsDebugState();
    const sheet1ImagesAfter = sheet1After.drawings.filter((d) => d.kind === "image");
    expect(sheet1ImagesAfter).toHaveLength(1);
    expect(sheet1ImagesAfter[0]!.id).toBe(inserted.id);
    expect(sheet1After.selectedId).toBe(null);

    app.destroy();
    root.remove();
  });

  it("clears drawing selection when restoreDocumentState changes the active sheet", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    deleteSeededCharts(app);
    const doc: any = app.getDocument();

    // Seed a drawing on Sheet1 and select it.
    doc.setSheetDrawings("Sheet1", [
      {
        id: 123,
        kind: { type: "shape", label: "Box" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(120), cy: pxToEmu(60) } },
        zOrder: 0,
      },
    ]);
    app.selectDrawingById(123);

    expect(app.getDrawingsDebugState().sheetId).toBe("Sheet1");
    expect(app.getDrawingsDebugState().selectedId).toBe(123);

    // Restore a snapshot that only contains Sheet2 (and also contains a drawing with the same id).
    // Even if the id exists on the new sheet, selection should be cleared when the active sheet changes.
    const snapshotDoc = new DocumentController();
    snapshotDoc.setSheetDrawings("Sheet2", [
      {
        id: 123,
        kind: { type: "shape", label: "Box2" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(80), cy: pxToEmu(40) } },
        zOrder: 0,
      },
    ]);
    const snapshot = snapshotDoc.encodeState();
    await app.restoreDocumentState(snapshot);

    const state = app.getDrawingsDebugState();
    expect(state.sheetId).toBe("Sheet2");
    expect(state.drawings).toHaveLength(1);
    expect(state.drawings[0]!.id).toBe(123);
    expect(state.selectedId).toBe(null);

    app.destroy();
    root.remove();
  });

  it("inserts pictures into the original sheet even if the user switches sheets mid-insert", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    deleteSeededCharts(app);
    const doc: any = app.getDocument();

    // Ensure Sheet2 exists so we can switch away while the image bytes are still loading.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    let resolveBytes: ((buf: ArrayBuffer) => void) | null = null;
    const arrayBuffer = new Promise<ArrayBuffer>((resolve) => {
      resolveBytes = resolve;
    });
    const file = {
      name: "slow.png",
      type: "image/png",
      size: 4,
      arrayBuffer: () => arrayBuffer,
    } as any as File;

    const insertPromise = app.insertPicturesFromFiles([file], { placeAt: { row: 0, col: 0 } });

    // Switch sheets while `insertPicturesFromFiles` is awaiting the file bytes.
    app.activateSheet("Sheet2");
    expect(app.getDrawingsDebugState().sheetId).toBe("Sheet2");

    resolveBytes?.(new Uint8Array([1, 2, 3, 4]).buffer);
    await insertPromise;

    // Picture should land on Sheet1 (where the insert was initiated), not the newly active sheet.
    const sheet1Drawings = Array.isArray(doc.getSheetDrawings("Sheet1")) ? doc.getSheetDrawings("Sheet1") : [];
    expect(sheet1Drawings).toHaveLength(1);
    expect(Array.isArray(doc.getSheetDrawings("Sheet2")) ? doc.getSheetDrawings("Sheet2") : []).toHaveLength(0);
    // And it should not disrupt the current sheet's drawing selection.
    expect(app.getDrawingsDebugState().selectedId).toBe(null);

    // WorkbookImageManager should track references even when drawings change on a non-active sheet.
    const imageId = (sheet1Drawings[0] as any)?.kind?.imageId as string | undefined;
    expect(typeof imageId).toBe("string");
    const imageManager = (app as any).workbookImageManager;
    expect(imageManager?.imageRefCount?.get?.(imageId!)).toBe(1);

    app.destroy();
    root.remove();
  });
});
