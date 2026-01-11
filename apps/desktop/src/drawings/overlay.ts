import type { Anchor, DrawingObject, ImageStore, Rect } from "./types";
import { ImageBitmapCache } from "./imageBitmapCache";

export const EMU_PER_INCH = 914_400;
export const PX_PER_INCH = 96;
export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH;

export interface GridGeometry {
  /** Sheet-space pixel origin for the top-left of a cell. */
  cellOriginPx(cell: { row: number; col: number }): { x: number; y: number };
  /** Pixel size of a cell. */
  cellSizePx(cell: { row: number; col: number }): { width: number; height: number };
}

export interface Viewport {
  scrollX: number;
  scrollY: number;
  width: number;
  height: number;
  dpr: number;
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return px * EMU_PER_PX;
}

export function anchorToRectPx(anchor: Anchor, geom: GridGeometry): Rect {
  switch (anchor.type) {
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        x: origin.x + emuToPx(anchor.from.offset.xEmu),
        y: origin.y + emuToPx(anchor.from.offset.yEmu),
        width: emuToPx(anchor.size.cx),
        height: emuToPx(anchor.size.cy),
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);

      const x1 = fromOrigin.x + emuToPx(anchor.from.offset.xEmu);
      const y1 = fromOrigin.y + emuToPx(anchor.from.offset.yEmu);

      // In DrawingML, `to` specifies the cell *containing* the bottom-right
      // corner (i.e. the first cell strictly outside the shape when the corner
      // lies on a grid boundary). The absolute end point is therefore the
      // origin of the `to` cell plus the offsets.
      const x2 = toOrigin.x + emuToPx(anchor.to.offset.xEmu);
      const y2 = toOrigin.y + emuToPx(anchor.to.offset.yEmu);

      return {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        width: Math.abs(x2 - x1),
        height: Math.abs(y2 - y1),
      };
    }
    case "absolute":
      return {
        x: emuToPx(anchor.pos.xEmu),
        y: emuToPx(anchor.pos.yEmu),
        width: emuToPx(anchor.size.cx),
        height: emuToPx(anchor.size.cy),
      };
  }
}

export class DrawingOverlay {
  private readonly ctx: CanvasRenderingContext2D;
  private readonly bitmapCache = new ImageBitmapCache();
  private selectedId: number | null = null;

  constructor(
    private readonly canvas: HTMLCanvasElement,
    private readonly images: ImageStore,
    private readonly geom: GridGeometry,
  ) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("overlay canvas 2d context not available");
    this.ctx = ctx;
  }

  resize(viewport: Viewport): void {
    this.canvas.width = Math.floor(viewport.width * viewport.dpr);
    this.canvas.height = Math.floor(viewport.height * viewport.dpr);
    this.canvas.style.width = `${viewport.width}px`;
    this.canvas.style.height = `${viewport.height}px`;
    this.ctx.setTransform(viewport.dpr, 0, 0, viewport.dpr, 0, 0);
  }

  async render(objects: DrawingObject[], viewport: Viewport): Promise<void> {
    const ctx = this.ctx;
    ctx.clearRect(0, 0, viewport.width, viewport.height);

    const ordered = [...objects].sort((a, b) => a.zOrder - b.zOrder);

    for (const obj of ordered) {
      const rect = anchorToRectPx(obj.anchor, this.geom);
      const screenRect = {
        x: rect.x - viewport.scrollX,
        y: rect.y - viewport.scrollY,
        width: rect.width,
        height: rect.height,
      };

      if (!intersects(screenRect, { x: 0, y: 0, width: viewport.width, height: viewport.height })) {
        continue;
      }

      if (obj.kind.type === "image") {
        const entry = this.images.get(obj.kind.imageId);
        if (!entry) continue;
        const bitmap = await this.bitmapCache.get(entry);
        ctx.drawImage(bitmap, screenRect.x, screenRect.y, screenRect.width, screenRect.height);
        continue;
      }

      // Placeholder rendering for shapes/charts/unknown.
      ctx.save();
      const chartStroke = resolveCssVar("--chart-series-1", "blue");
      const shapeStroke = resolveCssVar("--chart-series-2", "cyan");
      const labelColor = resolveCssVar("--text-primary", "black");

      ctx.strokeStyle = obj.kind.type === "chart" ? chartStroke : shapeStroke;
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 2]);
      ctx.strokeRect(screenRect.x, screenRect.y, screenRect.width, screenRect.height);
      ctx.setLineDash([]);
      ctx.fillStyle = labelColor;
      ctx.globalAlpha = 0.6;
      ctx.font = "12px sans-serif";
      ctx.fillText(obj.kind.type, screenRect.x + 4, screenRect.y + 14);
      ctx.restore();
    }

    // Selection overlay.
    if (this.selectedId != null) {
      const selected = objects.find((o) => o.id === this.selectedId);
      if (selected) {
        const rect = anchorToRectPx(selected.anchor, this.geom);
        const screen = {
          x: rect.x - viewport.scrollX,
          y: rect.y - viewport.scrollY,
          width: rect.width,
          height: rect.height,
        };
        drawSelection(ctx, screen);
      }
    }
  }

  /**
   * Optional selection id used for rendering handles.
   *
   * This is intentionally not part of the constructor API so the overlay can be
   * integrated incrementally with different state management approaches.
   */
  setSelectedId(id: number | null): void {
    this.selectedId = id;
  }
}

function intersects(a: Rect, b: Rect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function drawSelection(ctx: CanvasRenderingContext2D, rect: Rect): void {
  ctx.save();
  const borderColor = resolveCssVar("--selection-border", "blue");
  const handleBg = resolveCssVar("--bg-primary", "white");

  ctx.strokeStyle = borderColor;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);

  const handle = 8;
  const half = handle / 2;
  const points = [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x, y: rect.y + rect.height },
  ];

  ctx.fillStyle = handleBg;
  ctx.strokeStyle = borderColor;
  ctx.lineWidth = 1;
  for (const p of points) {
    ctx.beginPath();
    ctx.rect(p.x - half, p.y - half, handle, handle);
    ctx.fill();
    ctx.stroke();
  }
  ctx.restore();
}

function resolveCssVar(name: string, fallback: string): string {
  try {
    if (typeof document === "undefined") return fallback;
    const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return value.length ? value : fallback;
  } catch {
    return fallback;
  }
}
