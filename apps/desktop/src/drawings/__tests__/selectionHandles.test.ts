import { describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../overlay";
import { DrawingInteractionController } from "../interaction";
import {
  cursorForResizeHandle,
  cursorForResizeHandleWithTransform,
  getResizeHandleCenters,
  hitTestResizeHandle,
  type ResizeHandle,
} from "../selectionHandles";
import type { DrawingTransform } from "../types";
import type { DrawingObject } from "../types";
import type { GridGeometry, Viewport } from "../overlay";

function makePointerEvent(x: number, y: number, pointerId = 1): any {
  return {
    clientX: x,
    clientY: y,
    pointerId,
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
    stopImmediatePropagation: vi.fn(),
  };
}

function makeCanvas(listeners: Map<string, (e: any) => void>): any {
  return {
    style: { cursor: "" },
    addEventListener: (type: string, cb: (e: any) => void) => listeners.set(type, cb),
    removeEventListener: (type: string) => listeners.delete(type),
    setPointerCapture: vi.fn(),
    releasePointerCapture: vi.fn(),
    getBoundingClientRect: () => ({ left: 0, top: 0 }) as DOMRect,
  };
}

describe("drawings selection handles", () => {
  it("hitTestResizeHandle detects all 8 handles", () => {
    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    const centers = getResizeHandleCenters(bounds);
    for (const c of centers) {
      expect(hitTestResizeHandle(bounds, c.x, c.y)).toBe(c.handle);
    }

    // Center should not register as a handle.
    expect(hitTestResizeHandle(bounds, 140, 220)).toBeNull();
  });

  it("hitTestResizeHandle detects handles for rotated objects", () => {
    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    const transform: DrawingTransform = { rotationDeg: 90, flipH: false, flipV: false };

    const centers = getResizeHandleCenters(bounds, transform);
    for (const c of centers) {
      expect(hitTestResizeHandle(bounds, c.x, c.y, transform)).toBe(c.handle);
    }
  });

  it("cursorForResizeHandle matches expected CSS cursors", () => {
    const expected: Record<ResizeHandle, string> = {
      nw: "nwse-resize",
      n: "ns-resize",
      ne: "nesw-resize",
      e: "ew-resize",
      se: "nwse-resize",
      s: "ns-resize",
      sw: "nesw-resize",
      w: "ew-resize",
    };

    for (const [handle, cursor] of Object.entries(expected) as Array<[ResizeHandle, string]>) {
      expect(cursorForResizeHandle(handle)).toBe(cursor);
    }
  });

  it("cursorForResizeHandleWithTransform matches cursorForResizeHandle at rotationDeg=0", () => {
    const identity: DrawingTransform = { rotationDeg: 0, flipH: false, flipV: false };
    const handles: ResizeHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
    for (const handle of handles) {
      expect(cursorForResizeHandleWithTransform(handle, identity)).toBe(cursorForResizeHandle(handle));
    }
  });

  it("cursorForResizeHandleWithTransform accounts for object rotation", () => {
    const t90: DrawingTransform = { rotationDeg: 90, flipH: false, flipV: false };
    expect(cursorForResizeHandleWithTransform("n", t90)).toBe("ew-resize");
    expect(cursorForResizeHandleWithTransform("s", t90)).toBe("ew-resize");
    expect(cursorForResizeHandleWithTransform("e", t90)).toBe("ns-resize");
    expect(cursorForResizeHandleWithTransform("w", t90)).toBe("ns-resize");

    // 90° rotation swaps the diagonals.
    expect(cursorForResizeHandleWithTransform("nw", t90)).toBe("nesw-resize");
    expect(cursorForResizeHandleWithTransform("se", t90)).toBe("nesw-resize");
    expect(cursorForResizeHandleWithTransform("ne", t90)).toBe("nwse-resize");
    expect(cursorForResizeHandleWithTransform("sw", t90)).toBe("nwse-resize");
  });

  it("cursorForResizeHandleWithTransform is resilient to flips", () => {
    const handles: ResizeHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
    const cursors = new Set(["ns-resize", "ew-resize", "nwse-resize", "nesw-resize"]);
    const transforms: DrawingTransform[] = [
      { rotationDeg: 0, flipH: true, flipV: false },
      { rotationDeg: 0, flipH: false, flipV: true },
      { rotationDeg: 0, flipH: true, flipV: true },
      { rotationDeg: 45, flipH: true, flipV: false },
      { rotationDeg: 45, flipH: false, flipV: true },
      { rotationDeg: 45, flipH: true, flipV: true },
    ];
    for (const transform of transforms) {
      for (const handle of handles) {
        expect(cursors.has(cursorForResizeHandleWithTransform(handle, transform))).toBe(true);
      }
    }
  });

  it("DrawingInteractionController updates cursor on hover for selected object handles", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    const pointerMove = listeners.get("pointermove");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");
    expect(pointerMove).toBeTypeOf("function");

    // Click once to select (and finish the drag sequence).
    pointerDown!(makePointerEvent(140, 220));
    pointerUp!(makePointerEvent(140, 220));

    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    for (const c of getResizeHandleCenters(bounds)) {
      pointerMove!(makePointerEvent(c.x, c.y));
      expect(canvas.style.cursor).toBe(cursorForResizeHandle(c.handle));
    }
  });

  it("DrawingInteractionController hit-tests transformed handles for hover cursors", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    const transform = { rotationDeg: 90, flipH: false, flipV: false };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
        transform,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    const pointerMove = listeners.get("pointermove");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");
    expect(pointerMove).toBeTypeOf("function");

    // Select the object.
    pointerDown!(makePointerEvent(140, 220));
    pointerUp!(makePointerEvent(140, 220));

    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    for (const c of getResizeHandleCenters(bounds, transform)) {
      pointerMove!(makePointerEvent(c.x, c.y));
      expect(canvas.style.cursor).toBe(cursorForResizeHandleWithTransform(c.handle, transform));
    }
  });

  it("DrawingInteractionController starts resize from transformed handles (pointerdown)", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    const transform = { rotationDeg: 90, flipH: false, flipV: false };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
        transform,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");

    // Select the object.
    pointerDown!(makePointerEvent(140, 220, 1));
    pointerUp!(makePointerEvent(140, 220, 1));
    expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);

    // Each transformed handle center should be clickable to start a resize without changing selection.
    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    const centers = getResizeHandleCenters(bounds, transform);
    for (let i = 0; i < centers.length; i += 1) {
      const c = centers[i]!;
      const id = i + 2;
      pointerDown!(makePointerEvent(c.x, c.y, id));
      expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);
      expect(canvas.style.cursor).toBe(cursorForResizeHandle(c.handle, transform));
      pointerUp!(makePointerEvent(c.x, c.y, id));
    }
  });

  it("allows starting a resize from handle hit areas outside the object bounds", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    const initial: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
        size: { cx: pxToEmu(80), cy: pxToEmu(40) },
      },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [initial];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown");
    const pointerUp = listeners.get("pointerup");
    const pointerMove = listeners.get("pointermove");
    expect(pointerDown).toBeTypeOf("function");
    expect(pointerUp).toBeTypeOf("function");
    expect(pointerMove).toBeTypeOf("function");

    // Select the object first.
    pointerDown!(makePointerEvent(140, 220));
    pointerUp!(makePointerEvent(140, 220));
    expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);

    // Click slightly outside the object's bounds but within the top-left handle hit box.
    // Object bounds: x in [100, 180], y in [200, 240]. This point is outside (x<100,y<200).
    pointerDown!(makePointerEvent(97, 197));
    expect(callbacks.onSelectionChange).toHaveBeenCalledTimes(1);
    expect(canvas.style.cursor).toBe("nwse-resize");

    // Drag the handle up/left by 10px; this should expand the rect and move the origin.
    pointerMove!(makePointerEvent(87, 187));
    const resized = objects[0]!.anchor;
    expect(resized.type).toBe("absolute");
    if (resized.type !== "absolute") throw new Error("unexpected anchor type");
    expect(resized.pos.xEmu).toBeCloseTo(pxToEmu(90));
    expect(resized.pos.yEmu).toBeCloseTo(pxToEmu(190));
    expect(resized.size.cx).toBeCloseTo(pxToEmu(90));
    expect(resized.size.cy).toBeCloseTo(pxToEmu(50));

    pointerUp!(makePointerEvent(87, 187));
  });

  it("supports hover cursors with header offsets (shared-grid headers)", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = {
      scrollX: 0,
      scrollY: 0,
      width: 500,
      height: 500,
      dpr: 1,
      headerOffsetX: 10,
      headerOffsetY: 20,
    };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown")!;
    const pointerUp = listeners.get("pointerup")!;
    const pointerMove = listeners.get("pointermove")!;

    // Select.
    pointerDown(makePointerEvent(140 + viewport.headerOffsetX!, 220 + viewport.headerOffsetY!));
    pointerUp(makePointerEvent(140 + viewport.headerOffsetX!, 220 + viewport.headerOffsetY!));

    const bounds = {
      x: 100 + viewport.headerOffsetX!,
      y: 200 + viewport.headerOffsetY!,
      width: 80,
      height: 40,
    };
    for (const c of getResizeHandleCenters(bounds)) {
      pointerMove(makePointerEvent(c.x, c.y));
      expect(canvas.style.cursor).toBe(cursorForResizeHandle(c.handle));
    }
  });

  it("supports hover cursors in frozen panes when the sheet is scrolled", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    // Use a large frozen pane boundary so the entire object (and all 8 handles) are
    // contained within the frozen quadrant, even while the sheet is scrolled.
    const CELL = 200;
    const geom: GridGeometry = {
      cellOriginPx: (cell) => ({ x: cell.col * CELL, y: cell.row * CELL }),
      cellSizePx: () => ({ width: CELL, height: CELL }),
    };

    const viewport: Viewport = {
      scrollX: 50,
      scrollY: 100,
      width: 500,
      height: 500,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown")!;
    const pointerUp = listeners.get("pointerup")!;
    const pointerMove = listeners.get("pointermove")!;

    // Select (object is in the frozen top-left pane, so it's at (0,0) on screen).
    pointerDown(makePointerEvent(40, 20));
    pointerUp(makePointerEvent(40, 20));

    const bounds = { x: 0, y: 0, width: 80, height: 40 };
    for (const c of getResizeHandleCenters(bounds)) {
      pointerMove(makePointerEvent(c.x, c.y));
      expect(canvas.style.cursor).toBe(cursorForResizeHandle(c.handle));
    }
  });

  it("does not show resize/move cursors for a selected object outside its frozen quadrant", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const CELL = 10;
    const geom: GridGeometry = {
      cellOriginPx: (cell) => ({ x: cell.col * CELL, y: cell.row * CELL }),
      cellSizePx: () => ({ width: CELL, height: CELL }),
    };

    // Very small frozen quadrant so the object bounds extend past it.
    const viewport: Viewport = {
      scrollX: 0,
      scrollY: 0,
      width: 500,
      height: 500,
      dpr: 1,
      frozenRows: 1,
      frozenCols: 1,
      frozenWidthPx: CELL,
      frozenHeightPx: CELL,
    };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown")!;
    const pointerUp = listeners.get("pointerup")!;
    const pointerMove = listeners.get("pointermove")!;

    // Select the object by clicking in the visible frozen top-left quadrant.
    pointerDown(makePointerEvent(5, 5));
    pointerUp(makePointerEvent(5, 5));

    // Hover where the top-right handle would have been if the object wasn't clipped; this point is
    // in the top-right quadrant, so no resize cursor should be shown.
    pointerMove(makePointerEvent(80, 0));
    expect(canvas.style.cursor).toBe("default");

    // The point is also inside the object's *unclipped* rect, so ensure we don't incorrectly show
    // the move cursor either (cursor remains default).
    pointerMove(makePointerEvent(40, 5));
    expect(canvas.style.cursor).toBe("default");
  });

  it("resets the cursor to default on pointerleave when idle", () => {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = makeCanvas(listeners);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

    let objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape", label: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(80), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ];

    const callbacks = {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next: DrawingObject[]) => {
        objects = next;
      },
      onSelectionChange: vi.fn(),
    };

    new DrawingInteractionController(canvas as HTMLCanvasElement, geom, callbacks);

    const pointerDown = listeners.get("pointerdown")!;
    const pointerUp = listeners.get("pointerup")!;
    const pointerMove = listeners.get("pointermove")!;
    const pointerLeave = listeners.get("pointerleave")!;

    // Select then hover a handle.
    pointerDown(makePointerEvent(140, 220));
    pointerUp(makePointerEvent(140, 220));
    pointerMove(makePointerEvent(100, 200));
    expect(canvas.style.cursor).toBe("nwse-resize");

    // Leaving the canvas should reset the cursor.
    pointerLeave({});
    expect(canvas.style.cursor).toBe("default");
  });
});
