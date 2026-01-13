/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { DrawingInteractionController } from "../interaction";
import { anchorToRectPx, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject } from "../types";

type Listener = (e: any) => void;

class StubEventTarget {
  private readonly listeners = new Map<string, Array<{ listener: Listener; capture: boolean }>>();
  readonly style: { cursor?: string } = {};

  constructor(private readonly rect: { left: number; top: number; width?: number; height?: number }) {}

  addEventListener(type: string, listener: Listener, options?: boolean | AddEventListenerOptions): void {
    const capture = typeof options === "boolean" ? options : Boolean(options?.capture);
    const entries = this.listeners.get(type) ?? [];
    entries.push({ listener, capture });
    this.listeners.set(type, entries);
  }

  removeEventListener(type: string, listener: Listener, options?: boolean | EventListenerOptions): void {
    const capture = typeof options === "boolean" ? options : Boolean((options as AddEventListenerOptions | undefined)?.capture);
    const entries = this.listeners.get(type) ?? [];
    this.listeners.set(
      type,
      entries.filter((entry) => entry.listener !== listener || entry.capture !== capture),
    );
  }

  getBoundingClientRect(): DOMRect {
    const width = this.rect.width ?? 300;
    const height = this.rect.height ?? 200;
    const left = this.rect.left;
    const top = this.rect.top;
    return {
      left,
      top,
      right: left + width,
      bottom: top + height,
      width,
      height,
      x: left,
      y: top,
      toJSON: () => ({}),
    } as unknown as DOMRect;
  }

  setPointerCapture(): void {
    // no-op for tests
  }

  releasePointerCapture(): void {
    // no-op for tests
  }

  dispatchPointerEvent(type: string, e: any): void {
    const entries = this.listeners.get(type) ?? [];
    const capture = entries.filter((entry) => entry.capture);
    const bubble = entries.filter((entry) => !entry.capture);

    for (const entry of capture) {
      if (e._immediateStopped) return;
      entry.listener(e);
    }

    if (e._propagationStopped || e._immediateStopped) return;

    for (const entry of bubble) {
      if (e._immediateStopped) return;
      entry.listener(e);
    }
  }
}

function createPointerEvent(init: { clientX: number; clientY: number; pointerId: number }): any {
  return {
    ...init,
    defaultPrevented: false,
    _propagationStopped: false,
    _immediateStopped: false,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this._propagationStopped = true;
    },
    stopImmediatePropagation() {
      this._immediateStopped = true;
      this._propagationStopped = true;
    },
  };
}

function createGeom(opts: { colWidths: number[]; rowHeights: number[] }): GridGeometry {
  const { colWidths, rowHeights } = opts;

  const colPrefix = [0];
  for (let i = 0; i < colWidths.length; i += 1) {
    colPrefix[i + 1] = colPrefix[i]! + colWidths[i]!;
  }

  const rowPrefix = [0];
  for (let i = 0; i < rowHeights.length; i += 1) {
    rowPrefix[i + 1] = rowPrefix[i]! + rowHeights[i]!;
  }

  const sumBefore = (prefix: number[], idx: number): number => {
    if (idx <= 0) return 0;
    if (idx < prefix.length) return prefix[idx]!;
    // Extend using last known size.
    const lastSize = prefix.length >= 2 ? prefix[prefix.length - 1]! - prefix[prefix.length - 2]! : 0;
    return prefix[prefix.length - 1]! + (idx - (prefix.length - 1)) * lastSize;
  };

  const getSize = (sizes: number[], idx: number): number => {
    if (idx < 0) return 0;
    return sizes[idx] ?? sizes[sizes.length - 1] ?? 0;
  };

  return {
    cellOriginPx: ({ row, col }) => ({
      x: sumBefore(colPrefix, col),
      y: sumBefore(rowPrefix, row),
    }),
    cellSizePx: ({ row, col }) => ({
      width: getSize(colWidths, col),
      height: getSize(rowHeights, row),
    }),
  };
}

function createControllerHarness(opts: {
  geom: GridGeometry;
  viewport?: Partial<Viewport>;
  object: DrawingObject;
}): {
  controller: DrawingInteractionController;
  element: StubEventTarget;
  getObject(): DrawingObject;
  dispose(): void;
} {
  const el = new StubEventTarget({ left: 0, top: 0, width: 500, height: 500 });

  const viewport: Viewport = {
    scrollX: 0,
    scrollY: 0,
    width: 500,
    height: 500,
    dpr: 1,
    zoom: 1,
    ...opts.viewport,
  };

  let objects: DrawingObject[] = [opts.object];
  const controller = new DrawingInteractionController(el as unknown as HTMLElement, opts.geom, {
    getViewport: () => viewport,
    getObjects: () => objects,
    setObjects: (next) => {
      objects = next;
    },
  });

  return {
    controller,
    element: el,
    getObject: () => objects[0]!,
    dispose: () => controller.dispose(),
  };
}

describe("DrawingInteractionController anchor normalization", () => {
  it("increments from.cell.col when dragging right across multiple variable-width columns and keeps offsets bounded", () => {
    const geom = createGeom({ colWidths: [40, 60, 50, 50], rowHeights: [30] });
    const object: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "box" },
      zOrder: 0,
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
    };

    const { element, getObject, dispose } = createControllerHarness({ geom, object });
    try {
      // Start drag from the middle (avoid resize handles at corners).
      element.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 60, clientY: 40, pointerId: 1 }));
      // Drag right by 120px: col0(40) + col1(60) -> land in col2 with 20px offset.
      element.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 180, clientY: 40, pointerId: 1 }));
      element.dispatchPointerEvent("pointerup", createPointerEvent({ clientX: 180, clientY: 40, pointerId: 1 }));

      const updated = getObject();
      expect(updated.anchor.type).toBe("oneCell");
      if (updated.anchor.type !== "oneCell") return;

      expect(updated.anchor.from.cell.col).toBe(2);
      expect(updated.anchor.from.offset.xEmu).toBe(pxToEmu(20));

      const col2WidthEmu = pxToEmu(50);
      expect(updated.anchor.from.offset.xEmu).toBeGreaterThanOrEqual(0);
      expect(updated.anchor.from.offset.xEmu).toBeLessThan(col2WidthEmu);

      // Normalization should preserve the absolute position.
      const rect = anchorToRectPx(updated.anchor, geom);
      expect(rect.x).toBe(120);
    } finally {
      dispose();
    }
  });

  it("updates from.cell.col when moving an object from frozen to non-frozen columns so scroll rules can switch", () => {
    const geom = createGeom({ colWidths: [50, 50, 50], rowHeights: [30] });
    const object: DrawingObject = {
      id: 1,
      kind: { type: "shape", label: "box" },
      zOrder: 0,
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        size: { cx: pxToEmu(80), cy: pxToEmu(40) },
      },
    };

    const { element, getObject, dispose } = createControllerHarness({
      geom,
      object,
      viewport: { frozenCols: 1, scrollX: 100 },
    });
    try {
      const frozenCols = 1;
      const scrollX = 100;
      const effectiveScrollX = (col: number) => (col < frozenCols ? 0 : scrollX);

      const before = getObject();
      const beforeScroll = effectiveScrollX(before.anchor.type === "oneCell" ? before.anchor.from.cell.col : 0);
      expect(beforeScroll).toBe(0);

      // Drag into column 1 (non-frozen when frozenCols=1).
      element.dispatchPointerEvent("pointerdown", createPointerEvent({ clientX: 20, clientY: 20, pointerId: 1 }));
      element.dispatchPointerEvent("pointermove", createPointerEvent({ clientX: 80, clientY: 20, pointerId: 1 })); // +60px
      element.dispatchPointerEvent("pointerup", createPointerEvent({ clientX: 80, clientY: 20, pointerId: 1 }));

      const after = getObject();
      expect(after.anchor.type).toBe("oneCell");
      if (after.anchor.type !== "oneCell") return;

      expect(after.anchor.from.cell.col).toBe(1);
      const afterScroll = effectiveScrollX(after.anchor.from.cell.col);
      expect(afterScroll).toBe(scrollX);

      // Demonstrate how the scroll behavior would change in a frozen-pane layout
      // that keys off `from.cell.col`.
      const beforeRect = anchorToRectPx(before.anchor, geom);
      const afterRect = anchorToRectPx(after.anchor, geom);
      expect(beforeRect.x - beforeScroll).toBe(0);
      expect(afterRect.x - afterScroll).toBe(-40);
    } finally {
      dispose();
    }
  });
});
