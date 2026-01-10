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
}): ScrollbarThumb {
  const minThumbSize = options.minThumbSize ?? 24;
  const trackSize = Math.max(0, options.trackSize);
  const viewportSize = Math.max(0, options.viewportSize);
  const contentSize = Math.max(0, options.contentSize);
  const maxScroll = Math.max(0, contentSize - viewportSize);
  const scrollPos = Math.min(Math.max(0, options.scrollPos), maxScroll);

  if (trackSize === 0) return { size: 0, offset: 0 };
  if (contentSize === 0 || maxScroll === 0) return { size: trackSize, offset: 0 };

  const rawThumbSize = (viewportSize / contentSize) * trackSize;
  const thumbSize = Math.min(trackSize, Math.max(minThumbSize, rawThumbSize));
  const thumbTravel = Math.max(0, trackSize - thumbSize);
  const offset = thumbTravel === 0 ? 0 : (scrollPos / maxScroll) * thumbTravel;

  return { size: thumbSize, offset };
}

