import type { AnchorPoint, DrawingObject } from "./types";
import type { GridGeometry, Viewport } from "./overlay";
import { anchorToRectPx, emuToPx, pxToEmu } from "./overlay";
import { buildHitTestIndex, hitTestDrawings, hitTestDrawingsObject, type HitTestIndex } from "./hitTest";
import { cursorForResizeHandle, hitTestResizeHandle, type ResizeHandle } from "./selectionHandles";
import { extractXfrmOff, patchXfrmExt, patchXfrmOff } from "./drawingml/patch";

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
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom);
    const inHeader = e.offsetX < paneLayout.headerOffsetX || e.offsetY < paneLayout.headerOffsetY;
    const pointInFrozenCols = !inHeader && e.offsetX < paneLayout.frozenBoundaryX;
    const pointInFrozenRows = !inHeader && e.offsetY < paneLayout.frozenBoundaryY;

    // Allow grabbing a resize handle for the current selection even when the
    // pointer is slightly outside the object's bounds (handles are centered on
    // the outline and extend half their size beyond the rect).
    const selectedIndex = this.selectedId != null ? index.byId.get(this.selectedId) : undefined;
    const selectedObject = selectedIndex != null ? index.ordered[selectedIndex] : undefined;
    if (selectedObject && !inHeader) {
      const objectPane = resolveAnchorPane(selectedObject.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
      if (objectPane.inFrozenCols === pointInFrozenCols && objectPane.inFrozenRows === pointInFrozenRows) {
        const selectedBounds = objectToScreenRect(selectedObject, viewport, this.geom, index.bounds[selectedIndex!]);
        const handle = hitTestResizeHandle(selectedBounds, e.offsetX, e.offsetY);
        if (handle) {
          this.canvas.setPointerCapture(e.pointerId);
          this.resizing = {
            id: selectedObject.id,
            handle,
            startX: e.offsetX,
            startY: e.offsetY,
            startObjects: objects,
          };
          this.canvas.style.cursor = cursorForResizeHandle(handle);
          return;
        }
      }
    }

    const hit = hitTestDrawings(index, viewport, e.offsetX, e.offsetY, this.geom);
    this.selectedId = hit?.object.id ?? null;
    this.callbacks.onSelectionChange?.(this.selectedId);
    if (!hit) {
      this.canvas.style.cursor = "default";
      return;
    }

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
      this.canvas.style.cursor = cursorForResizeHandle(handle);
    } else {
      this.dragging = {
        id: hit.object.id,
        startX: e.offsetX,
        startY: e.offsetY,
        startObjects: objects,
      };
      this.canvas.style.cursor = "move";
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
      this.canvas.style.cursor = cursorForResizeHandle(this.resizing.handle);
      return;
    }

    if (this.dragging) {
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
      this.canvas.style.cursor = "move";
      return;
    }

    this.updateCursor(e.offsetX, e.offsetY);
  };

  private readonly onPointerUp = (e: PointerEvent) => {
    const dragging = this.dragging;
    const resizing = this.resizing;
    if (!dragging && !resizing) return;

    // Commit-time patching only: pointermove updates anchors for live previews,
    // while pointerup updates preserved DrawingML fragments (`rawXml`, `xlsx.pic_xml`)
    // so inner `<a:xfrm>` values (when present) stay consistent with the new anchor.
    const objects = this.callbacks.getObjects();

    const active = dragging ?? resizing;
    const startObj = active.startObjects.find((o) => o.id === active.id);
    const currentObj = objects.find((o) => o.id === active.id);

    if (startObj && currentObj) {
      const startRect = anchorToRectPx(startObj.anchor, this.geom);
      const endRect = anchorToRectPx(currentObj.anchor, this.geom);

      const dxEmu = pxToEmu(endRect.x - startRect.x);
      const dyEmu = pxToEmu(endRect.y - startRect.y);
      const cxEmu = pxToEmu(endRect.width);
      const cyEmu = pxToEmu(endRect.height);

      let patched = currentObj;
      if (resizing) {
        patched = patchDrawingXmlForResize(patched, cxEmu, cyEmu);
      }
      if (dxEmu !== 0 || dyEmu !== 0) {
        patched = patchDrawingXmlForMove(patched, dxEmu, dyEmu);
      }

      if (patched !== currentObj) {
        this.callbacks.setObjects(objects.map((obj) => (obj.id === active.id ? patched : obj)));
      }
    }

    this.dragging = null;
    this.resizing = null;
    this.canvas.releasePointerCapture(e.pointerId);
    this.updateCursor(e.offsetX, e.offsetY);
  };

  private updateCursor(x: number, y: number): void {
    const viewport = this.callbacks.getViewport();
    const objects = this.callbacks.getObjects();
    const index = this.getHitTestIndex(objects);
    const paneLayout = resolveViewportPaneLayout(viewport, this.geom);
    if (x < paneLayout.headerOffsetX || y < paneLayout.headerOffsetY) {
      this.canvas.style.cursor = "default";
      return;
    }
    const pointInFrozenCols = x < paneLayout.frozenBoundaryX;
    const pointInFrozenRows = y < paneLayout.frozenBoundaryY;

    if (this.selectedId != null) {
      const selectedIndex = index.byId.get(this.selectedId);
      if (selectedIndex != null) {
        const selected = index.ordered[selectedIndex]!;
        const selectedPane = resolveAnchorPane(selected.anchor, paneLayout.frozenRows, paneLayout.frozenCols);
        if (
          selectedPane.inFrozenCols === pointInFrozenCols &&
          selectedPane.inFrozenRows === pointInFrozenRows
        ) {
          const bounds = objectToScreenRect(selected, viewport, this.geom, index.bounds[selectedIndex]);
          const handle = hitTestResizeHandle(bounds, x, y);
          if (handle) {
            this.canvas.style.cursor = cursorForResizeHandle(handle);
            return;
          }
          if (pointInRect(x, y, bounds)) {
            this.canvas.style.cursor = "move";
            return;
          }
        }
      }
    }

    const hit = hitTestDrawingsObject(index, viewport, x, y, this.geom);
    if (hit) {
      this.canvas.style.cursor = "move";
      return;
    }

    this.canvas.style.cursor = "default";
  }

  private getHitTestIndex(objects: readonly DrawingObject[]): HitTestIndex {
    if (this.hitTestIndex && this.hitTestIndexObjects === objects) return this.hitTestIndex;
    const built = buildHitTestIndex(objects, this.geom);
    this.hitTestIndex = built;
    this.hitTestIndexObjects = objects;
    return built;
  }
}

function patchDrawingXmlForResize(obj: DrawingObject, cxEmu: number, cyEmu: number): DrawingObject {
  return patchDrawingInnerXml(obj, (xml) => patchXfrmExt(xml, cxEmu, cyEmu));
}

function patchDrawingXmlForMove(obj: DrawingObject, dxEmu: number, dyEmu: number): DrawingObject {
  // Only patch a:xfrm/a:off when the existing representation uses non-zero off
  // values. When off is already 0, we keep it at 0 and rely on anchors.
  return patchDrawingInnerXml(obj, (xml) => {
    const off = extractXfrmOff(xml);
    if (!off) return xml;
    if (off.xEmu === 0 && off.yEmu === 0) return xml;
    return patchXfrmOff(xml, off.xEmu + dxEmu, off.yEmu + dyEmu);
  });
}

function patchDrawingInnerXml(obj: DrawingObject, patch: (xml: string) => string): DrawingObject {
  if (obj.kind.type === "image") {
    const picXml = obj.preserved?.["xlsx.pic_xml"];
    if (typeof picXml !== "string") return obj;
    const patched = patch(picXml);
    if (patched === picXml) return obj;
    return {
      ...obj,
      preserved: {
        ...(obj.preserved ?? {}),
        "xlsx.pic_xml": patched,
      },
    };
  }

  const kindAny = obj.kind as any;
  const rawXml: unknown = kindAny.rawXml ?? kindAny.raw_xml;
  if (typeof rawXml !== "string") return obj;
  const patched = patch(rawXml);
  if (patched === rawXml) return obj;
  return {
    ...obj,
    kind: { ...kindAny, rawXml: patched, raw_xml: patched },
  };
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
    case "e":
      right += dxPx;
      break;
    case "w":
      left += dxPx;
      break;
    case "s":
      bottom += dyPx;
      break;
    case "n":
      top += dyPx;
      break;
  }

  const movesLeftEdge = handle === "nw" || handle === "w" || handle === "sw";
  const movesTopEdge = handle === "nw" || handle === "n" || handle === "ne";

  // Prevent negative widths/heights by clamping the moved edges against the
  // fixed ones. This keeps the opposite edge stationary.
  if (right < left) {
    if (movesLeftEdge) {
      left = right;
    } else {
      right = left;
    }
  }
  if (bottom < top) {
    if (movesTopEdge) {
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

function objectToScreenRect(
  obj: DrawingObject,
  viewport: Viewport,
  geom: GridGeometry,
  sheetRect?: { x: number; y: number; width: number; height: number },
) {
  const rect = sheetRect ?? anchorToRectPx(obj.anchor, geom);
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  const pane = resolveAnchorPane(obj.anchor, frozenRows, frozenCols);
  const scrollX = pane.inFrozenCols ? 0 : viewport.scrollX;
  const scrollY = pane.inFrozenRows ? 0 : viewport.scrollY;
  return {
    x: rect.x - scrollX + headerOffsetX,
    y: rect.y - scrollY + headerOffsetY,
    width: rect.width,
    height: rect.height,
  };
}

function pointInRect(
  x: number,
  y: number,
  rect: { x: number; y: number; width: number; height: number },
): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

type PaneLayout = {
  frozenRows: number;
  frozenCols: number;
  headerOffsetX: number;
  headerOffsetY: number;
  frozenBoundaryX: number;
  frozenBoundaryY: number;
};

const PANE_CELL_SCRATCH = { row: 0, col: 0 };

function clampNumber(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function resolveViewportPaneLayout(viewport: Viewport, geom: GridGeometry): PaneLayout {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  let frozenBoundaryX = headerOffsetX;
  let frozenBoundaryY = headerOffsetY;

  if (frozenCols > 0) {
    let raw = viewport.frozenWidthPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        PANE_CELL_SCRATCH.row = 0;
        PANE_CELL_SCRATCH.col = frozenCols;
        derived = geom.cellOriginPx(PANE_CELL_SCRATCH).x;
      } catch {
        derived = 0;
      }
      raw = headerOffsetX + derived;
    }
    frozenBoundaryX = clampNumber(raw as number, headerOffsetX, viewport.width);
  }

  if (frozenRows > 0) {
    let raw = viewport.frozenHeightPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        PANE_CELL_SCRATCH.row = frozenRows;
        PANE_CELL_SCRATCH.col = 0;
        derived = geom.cellOriginPx(PANE_CELL_SCRATCH).y;
      } catch {
        derived = 0;
      }
      raw = headerOffsetY + derived;
    }
    frozenBoundaryY = clampNumber(raw as number, headerOffsetY, viewport.height);
  }

  return { frozenRows, frozenCols, headerOffsetX, headerOffsetY, frozenBoundaryX, frozenBoundaryY };
}

function resolveAnchorPane(
  anchor: DrawingObject["anchor"],
  frozenRows: number,
  frozenCols: number,
): { inFrozenRows: boolean; inFrozenCols: boolean } {
  if (anchor.type === "absolute") return { inFrozenRows: false, inFrozenCols: false };
  return {
    inFrozenRows: anchor.from.cell.row < frozenRows,
    inFrozenCols: anchor.from.cell.col < frozenCols,
  };
}
