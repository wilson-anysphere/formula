export interface ScrollbarThumb {
  size: number;
  offset: number;
}

export function computeScrollbarThumb(options: {
  scrollPos: number;
  viewportSize: number;
  contentSize: number;
  trackSize: number;
  minThumbSize?: number;
  /**
   * Optional output object to populate. When provided, `computeScrollbarThumb` will mutate and
   * return this object instead of allocating a new one.
   *
   * This is useful in high-frequency paths (e.g. scroll handlers) to reduce GC pressure.
   */
  out?: ScrollbarThumb;
}): ScrollbarThumb {
  const minThumbSize = options.minThumbSize ?? 24;
  const trackSize = Math.max(0, options.trackSize);
  const viewportSize = Math.max(0, options.viewportSize);
  const contentSize = Math.max(0, options.contentSize);
  const maxScroll = Math.max(0, contentSize - viewportSize);
  const scrollPos = Math.min(Math.max(0, options.scrollPos), maxScroll);

  const out = options.out;
  if (trackSize === 0) {
    if (out) {
      out.size = 0;
      out.offset = 0;
      return out;
    }
    return { size: 0, offset: 0 };
  }
  if (contentSize === 0 || maxScroll === 0) {
    if (out) {
      out.size = trackSize;
      out.offset = 0;
      return out;
    }
    return { size: trackSize, offset: 0 };
  }

  const rawThumbSize = (viewportSize / contentSize) * trackSize;
  const thumbSize = Math.min(trackSize, Math.max(minThumbSize, rawThumbSize));
  const thumbTravel = Math.max(0, trackSize - thumbSize);
  const offset = thumbTravel === 0 ? 0 : (scrollPos / maxScroll) * thumbTravel;

  if (out) {
    out.size = thumbSize;
    out.offset = offset;
    return out;
  }

  return { size: thumbSize, offset };
}
