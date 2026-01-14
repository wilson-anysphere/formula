/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../../drawings/modelAdapters";
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

describe("SpreadsheetApp drawing interaction commits", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

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

  it("persists drawing anchor/transform/preserved updates to DocumentController via onInteractionCommit (undoable)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", label: "Box", rawXml: "<before/>", raw_xml: "<before/>" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
      preserved: { foo: "before" },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    expect(before.anchor.type).toBe("absolute");
    if (before.anchor.type !== "absolute") {
      throw new Error("Expected absolute anchor for test drawing");
    }
    const after = {
      ...before,
      anchor: {
        ...before.anchor,
        // Move it slightly and keep the same size.
        pos: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) },
      },
      kind: { ...(before.kind as any), rawXml: "<after/>", raw_xml: "<after/>" },
      transform: { rotationDeg: 45, flipH: false, flipV: false },
      preserved: { foo: "after" },
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    expect(callbacks?.onInteractionCommit).toBeTypeOf("function");

    callbacks.onInteractionCommit({ kind: "rotate", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.id).toBe("drawing_foo");
    expect(updated?.zOrder).toBe(0);
    expect(updated?.anchor).toEqual(after.anchor);
    expect(updated?.kind?.rawXml).toBe("<after/>");
    expect(updated?.kind?.raw_xml).toBe("<after/>");
    expect(updated?.transform).toEqual(after.transform);
    expect(updated?.preserved).toEqual(after.preserved);

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
      expect(reverted?.anchor).toEqual(rawDrawing.anchor);
      expect(reverted?.kind?.rawXml).toBe("<before/>");
      expect(reverted?.kind?.raw_xml).toBe("<before/>");
      expect(reverted?.transform).toBeUndefined();
      expect(reverted?.preserved).toEqual(rawDrawing.preserved);
    }

    app.dispose();
    root.remove();
  });

  it("updates top-level size when present so resizes persist across adapter re-reads", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const startCx = pxToEmu(120);
    const startCy = pxToEmu(80);
    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: startCx, cy: startCy },
      },
      // Explicit size field (common for picture inserts).
      size: { cx: startCx, cy: startCy },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    expect(before.anchor.type).toBe("absolute");
    if (before.anchor.type !== "absolute") {
      throw new Error("Expected absolute anchor for test drawing");
    }

    // Simulate a resize: update anchor.size but leave `size` stale (DrawingInteractionController
    // updates anchors during resize but does not update the optional `size` field).
    const after = {
      ...before,
      anchor: {
        ...before.anchor,
        size: { cx: pxToEmu(200), cy: pxToEmu(150) },
      },
      // NOTE: intentionally do not update `size` here.
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    callbacks.onInteractionCommit({ kind: "resize", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.anchor?.size).toEqual({ cx: pxToEmu(200), cy: pxToEmu(150) });
    expect(updated?.size).toEqual({ cx: pxToEmu(200), cy: pxToEmu(150) });

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
      expect(reverted?.anchor?.size).toEqual({ cx: startCx, cy: startCy });
      expect(reverted?.size).toEqual({ cx: startCx, cy: startCy });
    }

    app.dispose();
    root.remove();
  });

  it("persists kind.rawXml patches from interaction commits (undoable)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", rawXml: "<before/>" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    const after = {
      ...before,
      kind: { ...(before.kind as any), rawXml: "<after/>" },
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    callbacks.onInteractionCommit({ kind: "move", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.kind?.rawXml ?? updated?.kind?.raw_xml).toBe("<after/>");

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
      expect(reverted?.kind?.rawXml ?? reverted?.kind?.raw_xml).toBe("<before/>");
    }

    app.dispose();
    root.remove();
  });

  it("persists kind.rawXml patches for externally-tagged kind encodings without corrupting the enum shape", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      // Externally-tagged enum (Rust model encoding).
      kind: { Shape: { raw_xml: "<before/>" } },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    const after = {
      ...before,
      kind: { ...(before.kind as any), rawXml: "<after/>" },
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    callbacks.onInteractionCommit({ kind: "move", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.kind).toBeTruthy();
    expect(Object.keys(updated.kind ?? {})).toEqual(["Shape"]);
    expect(updated?.kind?.Shape?.raw_xml ?? updated?.kind?.Shape?.rawXml).toBe("<after/>");

    if (typeof doc.undo === "function") {
      expect(doc.undo()).toBe(true);
      const reverted = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
      expect(Object.keys(reverted?.kind ?? {})).toEqual(["Shape"]);
      expect(reverted?.kind?.Shape?.raw_xml ?? reverted?.kind?.Shape?.rawXml).toBe("<before/>");
    }

    app.dispose();
    root.remove();
  });

  it("ignores drawing interaction commits when the app is read-only", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    expect(before.anchor.type).toBe("absolute");
    if (before.anchor.type !== "absolute") {
      throw new Error("Expected absolute anchor for test drawing");
    }
    const after = {
      ...before,
      anchor: {
        ...before.anchor,
        pos: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) },
      },
    };

    // Simulate read-only session.
    (app as any).collabSession = { isReadOnly: () => true };

    const callbacks = (app as any).drawingInteractionCallbacks;
    callbacks.onInteractionCommit({ kind: "move", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.anchor).toEqual(rawDrawing.anchor);

    app.dispose();
    root.remove();
  });

  it("does not wipe malformed raw transform payloads on move commits", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;

    const rawDrawing = {
      id: "drawing_foo",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
      // Malformed transform: missing flipH/flipV (adapter should ignore it).
      transform: { rotationDeg: 30 },
    };
    doc.setSheetDrawings(sheetId, [rawDrawing]);

    const before = convertDocumentSheetDrawingsToUiDrawingObjects(doc.getSheetDrawings(sheetId), { sheetId })[0]!;
    expect(before.transform).toBeUndefined();
    expect(before.anchor.type).toBe("absolute");
    if (before.anchor.type !== "absolute") {
      throw new Error("Expected absolute anchor for test drawing");
    }

    const after = {
      ...before,
      anchor: {
        ...before.anchor,
        pos: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) },
      },
    };

    const callbacks = (app as any).drawingInteractionCallbacks;
    callbacks.onInteractionCommit({ kind: "move", id: before.id, before, after, objects: [after] });

    const updated = doc.getSheetDrawings(sheetId).find((d: any) => String(d?.id) === "drawing_foo");
    expect(updated?.transform).toEqual(rawDrawing.transform);

    app.dispose();
    root.remove();
  });
});
