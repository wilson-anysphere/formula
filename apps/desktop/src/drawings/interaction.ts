import type { AnchorPoint, DrawingObject } from "./types";
import type { GridGeometry, Viewport } from "./overlay";
import { emuToPx, pxToEmu } from "./overlay";
import { buildHitTestIndex, hitTestDrawings, type HitTestIndex } from "./hitTest";

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
  private hitTestIndex: HitTestIndex | null = null;
  private hitTestIndexObjects: readonly DrawingObject[] | null = null;
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
    const index = this.getHitTestIndex(objects);
    const hit = hitTestDrawings(index, viewport, e.offsetX, e.offsetY);
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

      const next = this.resizing.startObjects.map((obj) => {
        if (obj.id !== this.resizing!.id) return obj;
        return {
          ...obj,
          anchor: resizeAnchor(obj.anchor, this.resizing!.handle, dx, dy, this.geom),
        };
      });
      this.callbacks.setObjects(next);
      return;
    }

    if (!this.dragging) return;
    const dx = e.offsetX - this.dragging.startX;
    const dy = e.offsetY - this.dragging.startY;

    const next = this.dragging.startObjects.map((obj) => {
      if (obj.id !== this.dragging!.id) return obj;
      return {
        ...obj,
        anchor: shiftAnchor(obj.anchor, dx, dy, this.geom),
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

  private getHitTestIndex(objects: readonly DrawingObject[]): HitTestIndex {
    if (this.hitTestIndex && this.hitTestIndexObjects === objects) return this.hitTestIndex;
    const built = buildHitTestIndex(objects, this.geom);
    this.hitTestIndex = built;
    this.hitTestIndexObjects = objects;
    return built;
  }
}

export function shiftAnchor(
  anchor: DrawingObject["anchor"],
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
): DrawingObject["anchor"] {
  switch (anchor.type) {
    case "oneCell":
      return {
        ...anchor,
        from: shiftAnchorPoint(anchor.from, dxPx, dyPx, geom),
      };
    case "twoCell":
      return {
        ...anchor,
        from: shiftAnchorPoint(anchor.from, dxPx, dyPx, geom),
        to: shiftAnchorPoint(anchor.to, dxPx, dyPx, geom),
      };
    case "absolute":
      return {
        ...anchor,
        pos: {
          xEmu: anchor.pos.xEmu + pxToEmu(dxPx),
          yEmu: anchor.pos.yEmu + pxToEmu(dyPx),
        },
      };
  }
}

export type ResizeHandle = "nw" | "ne" | "se" | "sw";

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

export function resizeAnchor(
  anchor: DrawingObject["anchor"],
  handle: ResizeHandle,
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
): DrawingObject["anchor"] {
  const rect =
    anchor.type === "absolute"
      ? {
          left: emuToPx(anchor.pos.xEmu),
          top: emuToPx(anchor.pos.yEmu),
          right: emuToPx(anchor.pos.xEmu + anchor.size.cx),
          bottom: emuToPx(anchor.pos.yEmu + anchor.size.cy),
        }
      : anchor.type === "oneCell"
        ? (() => {
            const p = anchorPointToSheetPx(anchor.from, geom);
            return {
              left: p.x,
              top: p.y,
              right: p.x + emuToPx(anchor.size.cx),
              bottom: p.y + emuToPx(anchor.size.cy),
            };
          })()
        : (() => {
            const from = anchorPointToSheetPx(anchor.from, geom);
            const to = anchorPointToSheetPx(anchor.to, geom);
            return { left: from.x, top: from.y, right: to.x, bottom: to.y };
          })();

  let { left, top, right, bottom } = rect;

  switch (handle) {
    case "se":
      right += dxPx;
      bottom += dyPx;
      break;
    case "nw":
      left += dxPx;
      top += dyPx;
      break;
    case "ne":
      right += dxPx;
      top += dyPx;
      break;
    case "sw":
      left += dxPx;
      bottom += dyPx;
      break;
  }

  // Prevent negative widths/heights by clamping the moved edges against the
  // fixed ones. This keeps the opposite corner stationary.
  if (right < left) {
    if (handle === "nw" || handle === "sw") {
      left = right;
    } else {
      right = left;
    }
  }
  if (bottom < top) {
    if (handle === "nw" || handle === "ne") {
      top = bottom;
    } else {
      bottom = top;
    }
  }

  const widthPx = Math.max(0, right - left);
  const heightPx = Math.max(0, bottom - top);

  switch (anchor.type) {
    case "oneCell": {
      const start = anchorPointToSheetPx(anchor.from, geom);
      const nextFrom = shiftAnchorPoint(anchor.from, left - start.x, top - start.y, geom);
      return {
        ...anchor,
        from: nextFrom,
        size: { cx: pxToEmu(widthPx), cy: pxToEmu(heightPx) },
      };
    }
    case "absolute": {
      return {
        ...anchor,
        pos: { xEmu: pxToEmu(left), yEmu: pxToEmu(top) },
        size: { cx: pxToEmu(widthPx), cy: pxToEmu(heightPx) },
      };
    }
    case "twoCell": {
      const startFrom = anchorPointToSheetPx(anchor.from, geom);
      const startTo = anchorPointToSheetPx(anchor.to, geom);
      const nextFrom = shiftAnchorPoint(anchor.from, left - startFrom.x, top - startFrom.y, geom);
      const nextTo = shiftAnchorPoint(anchor.to, right - startTo.x, bottom - startTo.y, geom);
      return { ...anchor, from: nextFrom, to: nextTo };
    }
  }
}

function anchorPointToSheetPx(point: AnchorPoint, geom: GridGeometry): { x: number; y: number } {
  const origin = geom.cellOriginPx(point.cell);
  return { x: origin.x + emuToPx(point.offset.xEmu), y: origin.y + emuToPx(point.offset.yEmu) };
}

const MAX_CELL_STEPS = 10_000;

export function shiftAnchorPoint(
  point: AnchorPoint,
  dxPx: number,
  dyPx: number,
  geom: GridGeometry,
): AnchorPoint {
  let row = point.cell.row;
  let col = point.cell.col;
  let xPx = emuToPx(point.offset.xEmu) + dxPx;
  let yPx = emuToPx(point.offset.yEmu) + dyPx;

  // Normalize X across column boundaries.
  for (let i = 0; i < MAX_CELL_STEPS && xPx < 0; i++) {
    if (col <= 0) {
      col = 0;
      xPx = 0;
      break;
    }
    col -= 1;
    const w = geom.cellSizePx({ row, col }).width;
    if (w <= 0) {
      xPx = 0;
      break;
    }
    xPx += w;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    const w = geom.cellSizePx({ row, col }).width;
    if (w <= 0) {
      xPx = 0;
      break;
    }
    if (xPx < w) break;
    xPx -= w;
    col += 1;
  }

  // Normalize Y across row boundaries.
  for (let i = 0; i < MAX_CELL_STEPS && yPx < 0; i++) {
    if (row <= 0) {
      row = 0;
      yPx = 0;
      break;
    }
    row -= 1;
    const h = geom.cellSizePx({ row, col }).height;
    if (h <= 0) {
      yPx = 0;
      break;
    }
    yPx += h;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    const h = geom.cellSizePx({ row, col }).height;
    if (h <= 0) {
      yPx = 0;
      break;
    }
    if (yPx < h) break;
    yPx -= h;
    row += 1;
  }

  // Best-effort clamp to avoid tiny float drift.
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    const w = geom.cellSizePx({ row, col }).width;
    if (w <= 0) {
      xPx = 0;
      break;
    }
    if (xPx < 0) xPx = 0;
    if (xPx < w) break;
    xPx -= w;
    col += 1;
  }
  for (let i = 0; i < MAX_CELL_STEPS; i++) {
    const h = geom.cellSizePx({ row, col }).height;
    if (h <= 0) {
      yPx = 0;
      break;
    }
    if (yPx < 0) yPx = 0;
    if (yPx < h) break;
    yPx -= h;
    row += 1;
  }

  return {
    ...point,
    cell: { row, col },
    offset: { xEmu: pxToEmu(xPx), yEmu: pxToEmu(yPx) },
  };
}
