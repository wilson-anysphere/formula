export interface WheelDeltaToPixelsOptions {
  /**
   * Approximate pixel height of a "line" scroll when `WheelEvent.deltaMode === DOM_DELTA_LINE`.
   *
   * Browsers disagree on line-mode deltas (Firefox typically emits `deltaY=3`),
   * so we convert the line unit into pixels to keep scrolling speed consistent.
   */
  lineHeight?: number;
  /**
   * Page size (in CSS px) used when `WheelEvent.deltaMode === DOM_DELTA_PAGE`.
   *
   * Callers should generally pass the viewport width/height for the axis being scrolled.
   */
  pageSize?: number;
}

export function wheelDeltaToPixels(delta: number, deltaMode: number, options?: WheelDeltaToPixelsOptions): number {
  if (!Number.isFinite(delta)) return 0;

  switch (deltaMode) {
    // DOM_DELTA_PIXEL
    case 0:
      return delta;
    // DOM_DELTA_LINE
    case 1: {
      const lineHeight = options?.lineHeight ?? 16;
      return delta * lineHeight;
    }
    // DOM_DELTA_PAGE
    case 2: {
      const pageSize = options?.pageSize ?? 800;
      return delta * pageSize;
    }
    default:
      return delta;
  }
}

