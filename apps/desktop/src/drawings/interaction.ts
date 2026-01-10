import type { DrawingObject } from "./types";
import type { GridGeometry, Viewport } from "./overlay";
import { pxToEmu } from "./overlay";
import { hitTestDrawings } from "./hitTest";

export interface DrawingInteractionCallbacks {
  getViewport(): Viewport;
  getObjects(): DrawingObject[];
  setObjects(next: DrawingObject[]): void;
  onSelectionChange?(selectedId: number | null): void;
}

/**
 * Minimal MVP interactions: click-to-select and drag to move.
 */
export class DrawingInteractionController {
  private dragging:
    | { id: number; startX: number; startY: number; startObjects: DrawingObject[] }
    | null = null;
  private resizing:
    | {
        id: number;
        handle: ResizeHandle;
        startX: number;
        startY: number;
        startObjects: DrawingObject[];
      }
    | null = null;
  private selectedId: number | null = null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly geom: GridGeometry,
    private readonly callbacks: DrawingInteractionCallbacks,
  ) {
    this.canvas.addEventListener("pointerdown", this.onPointerDown);
    this.canvas.addEventListener("pointermove", this.onPointerMove);
    this.canvas.addEventListener("pointerup", this.onPointerUp);
    this.canvas.addEventListener("pointercancel", this.onPointerUp);
  }

  dispose(): void {
    this.canvas.removeEventListener("pointerdown", this.onPointerDown);
    this.canvas.removeEventListener("pointermove", this.onPointerMove);
    this.canvas.removeEventListener("pointerup", this.onPointerUp);
    this.canvas.removeEventListener("pointercancel", this.onPointerUp);
  }

  private readonly onPointerDown = (e: PointerEvent) => {
    const viewport = this.callbacks.getViewport();
    const objects = this.callbacks.getObjects();
    const hit = hitTestDrawings(objects, viewport, this.geom, e.offsetX, e.offsetY);
    this.selectedId = hit?.object.id ?? null;
    this.callbacks.onSelectionChange?.(this.selectedId);
    if (!hit) return;

    this.canvas.setPointerCapture(e.pointerId);
    const handle = hitTestResizeHandle(hit.bounds, e.offsetX, e.offsetY);
    if (handle) {
      this.resizing = {
        id: hit.object.id,
        handle,
        startX: e.offsetX,
        startY: e.offsetY,
        startObjects: objects,
      };
    } else {
      this.dragging = {
        id: hit.object.id,
        startX: e.offsetX,
        startY: e.offsetY,
        startObjects: objects,
      };
    }
  };

  private readonly onPointerMove = (e: PointerEvent) => {
    if (this.resizing) {
      const dx = e.offsetX - this.resizing.startX;
      const dy = e.offsetY - this.resizing.startY;
      const dxEmu = pxToEmu(dx);
      const dyEmu = pxToEmu(dy);

      const next = this.resizing.startObjects.map((obj) => {
        if (obj.id !== this.resizing!.id) return obj;
        return {
          ...obj,
          anchor: resizeAnchor(obj.anchor, this.resizing!.handle, dxEmu, dyEmu),
        };
      });
      this.callbacks.setObjects(next);
      return;
    }

    if (!this.dragging) return;
    const dx = e.offsetX - this.dragging.startX;
    const dy = e.offsetY - this.dragging.startY;
    const dxEmu = pxToEmu(dx);
    const dyEmu = pxToEmu(dy);

    const next = this.dragging.startObjects.map((obj) => {
      if (obj.id !== this.dragging!.id) return obj;
      return {
        ...obj,
        anchor: shiftAnchor(obj.anchor, dxEmu, dyEmu),
      };
    });
    this.callbacks.setObjects(next);
  };

  private readonly onPointerUp = (e: PointerEvent) => {
    if (!this.dragging && !this.resizing) return;
    this.dragging = null;
    this.resizing = null;
    this.canvas.releasePointerCapture(e.pointerId);
  };
}

function shiftAnchor(anchor: DrawingObject["anchor"], dxEmu: number, dyEmu: number): DrawingObject["anchor"] {
  switch (anchor.type) {
    case "oneCell":
      return {
        ...anchor,
        from: {
          ...anchor.from,
          offset: {
            xEmu: anchor.from.offset.xEmu + dxEmu,
            yEmu: anchor.from.offset.yEmu + dyEmu,
          },
        },
      };
    case "twoCell":
      return {
        ...anchor,
        from: {
          ...anchor.from,
          offset: {
            xEmu: anchor.from.offset.xEmu + dxEmu,
            yEmu: anchor.from.offset.yEmu + dyEmu,
          },
        },
        to: {
          ...anchor.to,
          offset: {
            xEmu: anchor.to.offset.xEmu + dxEmu,
            yEmu: anchor.to.offset.yEmu + dyEmu,
          },
        },
      };
    case "absolute":
      return {
        ...anchor,
        pos: {
          xEmu: anchor.pos.xEmu + dxEmu,
          yEmu: anchor.pos.yEmu + dyEmu,
        },
      };
  }
}

type ResizeHandle = "nw" | "ne" | "se" | "sw";

function hitTestResizeHandle(bounds: { x: number; y: number; width: number; height: number }, x: number, y: number): ResizeHandle | null {
  const size = 10;
  const half = size / 2;
  const corners: Array<{ handle: ResizeHandle; cx: number; cy: number }> = [
    { handle: "nw", cx: bounds.x, cy: bounds.y },
    { handle: "ne", cx: bounds.x + bounds.width, cy: bounds.y },
    { handle: "se", cx: bounds.x + bounds.width, cy: bounds.y + bounds.height },
    { handle: "sw", cx: bounds.x, cy: bounds.y + bounds.height },
  ];
  for (const c of corners) {
    if (
      x >= c.cx - half &&
      x <= c.cx + half &&
      y >= c.cy - half &&
      y <= c.cy + half
    ) {
      return c.handle;
    }
  }
  return null;
}

function resizeAnchor(anchor: DrawingObject["anchor"], handle: ResizeHandle, dxEmu: number, dyEmu: number): DrawingObject["anchor"] {
  const clamp = (n: number) => Math.max(0, n);

  switch (anchor.type) {
    case "oneCell": {
      switch (handle) {
        case "se":
          return {
            ...anchor,
            size: { cx: clamp(anchor.size.cx + dxEmu), cy: clamp(anchor.size.cy + dyEmu) },
          };
        case "nw":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu + dxEmu,
                yEmu: anchor.from.offset.yEmu + dyEmu,
              },
            },
            size: { cx: clamp(anchor.size.cx - dxEmu), cy: clamp(anchor.size.cy - dyEmu) },
          };
        case "ne":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu,
                yEmu: anchor.from.offset.yEmu + dyEmu,
              },
            },
            size: { cx: clamp(anchor.size.cx + dxEmu), cy: clamp(anchor.size.cy - dyEmu) },
          };
        case "sw":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu + dxEmu,
                yEmu: anchor.from.offset.yEmu,
              },
            },
            size: { cx: clamp(anchor.size.cx - dxEmu), cy: clamp(anchor.size.cy + dyEmu) },
          };
      }
    }
    case "absolute": {
      switch (handle) {
        case "se":
          return {
            ...anchor,
            size: { cx: clamp(anchor.size.cx + dxEmu), cy: clamp(anchor.size.cy + dyEmu) },
          };
        case "nw":
          return {
            ...anchor,
            pos: { xEmu: anchor.pos.xEmu + dxEmu, yEmu: anchor.pos.yEmu + dyEmu },
            size: { cx: clamp(anchor.size.cx - dxEmu), cy: clamp(anchor.size.cy - dyEmu) },
          };
        case "ne":
          return {
            ...anchor,
            pos: { xEmu: anchor.pos.xEmu, yEmu: anchor.pos.yEmu + dyEmu },
            size: { cx: clamp(anchor.size.cx + dxEmu), cy: clamp(anchor.size.cy - dyEmu) },
          };
        case "sw":
          return {
            ...anchor,
            pos: { xEmu: anchor.pos.xEmu + dxEmu, yEmu: anchor.pos.yEmu },
            size: { cx: clamp(anchor.size.cx - dxEmu), cy: clamp(anchor.size.cy + dyEmu) },
          };
      }
    }
    case "twoCell": {
      // Best-effort resizing by shifting the relevant anchor points.
      switch (handle) {
        case "se":
          return {
            ...anchor,
            to: {
              ...anchor.to,
              offset: {
                xEmu: anchor.to.offset.xEmu + dxEmu,
                yEmu: anchor.to.offset.yEmu + dyEmu,
              },
            },
          };
        case "nw":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu + dxEmu,
                yEmu: anchor.from.offset.yEmu + dyEmu,
              },
            },
          };
        case "ne":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu,
                yEmu: anchor.from.offset.yEmu + dyEmu,
              },
            },
            to: {
              ...anchor.to,
              offset: {
                xEmu: anchor.to.offset.xEmu + dxEmu,
                yEmu: anchor.to.offset.yEmu,
              },
            },
          };
        case "sw":
          return {
            ...anchor,
            from: {
              ...anchor.from,
              offset: {
                xEmu: anchor.from.offset.xEmu + dxEmu,
                yEmu: anchor.from.offset.yEmu,
              },
            },
            to: {
              ...anchor.to,
              offset: {
                xEmu: anchor.to.offset.xEmu,
                yEmu: anchor.to.offset.yEmu + dyEmu,
              },
            },
          };
      }
    }
  }
}
