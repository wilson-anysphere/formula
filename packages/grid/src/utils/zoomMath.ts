import type { GridViewportState } from "../virtualization/VirtualScrollManager";

// Keep zoom bounds aligned with the desktop layout pane zoom clamp so the app can
// persist values without "snapping" on reload.
export const MIN_GRID_ZOOM = 0.25;
export const MAX_GRID_ZOOM = 4.0;

export function clampZoom(value: number, min = MIN_GRID_ZOOM, max = MAX_GRID_ZOOM): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(max, Math.max(min, value));
}

export function computeZoomFromPinchDistance(options: {
  startZoom: number;
  startDistance: number;
  currentDistance: number;
  minZoom?: number;
  maxZoom?: number;
}): number {
  const minZoom = options.minZoom ?? MIN_GRID_ZOOM;
  const maxZoom = options.maxZoom ?? MAX_GRID_ZOOM;
  const startZoom = options.startZoom;

  if (!Number.isFinite(startZoom) || startZoom <= 0) return clampZoom(1, minZoom, maxZoom);
  if (!Number.isFinite(options.startDistance) || options.startDistance <= 0) return clampZoom(startZoom, minZoom, maxZoom);
  if (!Number.isFinite(options.currentDistance) || options.currentDistance <= 0) return clampZoom(startZoom, minZoom, maxZoom);

  const scale = options.currentDistance / options.startDistance;
  return clampZoom(startZoom * scale, minZoom, maxZoom);
}

export function computeAnchoredScrollAfterZoom(options: {
  viewport: Pick<GridViewportState, "width" | "height" | "frozenWidth" | "frozenHeight">;
  startZoom: number;
  nextZoom: number;
  startScroll: { x: number; y: number };
  /**
   * The anchor point in viewport coordinates at the start of the zoom gesture.
   *
   * For a static anchor (programmatic zoom), use the same point for start and next.
   */
  startAnchor: { x: number; y: number };
  /**
   * The anchor point in viewport coordinates after the zoom gesture update.
   *
   * When pinch fingers translate, this will differ from `startAnchor` to add pan.
   */
  nextAnchor: { x: number; y: number };
}): { x: number; y: number } {
  const { viewport, startZoom, nextZoom, startScroll, startAnchor, nextAnchor } = options;

  const ratio = startZoom === 0 ? 1 : nextZoom / startZoom;
  if (!Number.isFinite(ratio) || ratio <= 0) return { ...startScroll };

  const frozenWidth = Math.min(Math.max(0, viewport.frozenWidth), Math.max(0, viewport.width));
  const frozenHeight = Math.min(Math.max(0, viewport.frozenHeight), Math.max(0, viewport.height));

  const hasScrollableX = viewport.width > frozenWidth;
  const hasScrollableY = viewport.height > frozenHeight;

  const startInScrollableX = hasScrollableX && startAnchor.x >= frozenWidth;
  const startInScrollableY = hasScrollableY && startAnchor.y >= frozenHeight;

  const x = startInScrollableX ? (startScroll.x + startAnchor.x) * ratio - nextAnchor.x : startScroll.x * ratio;
  const y = startInScrollableY ? (startScroll.y + startAnchor.y) * ratio - nextAnchor.y : startScroll.y * ratio;

  return { x, y };
}
