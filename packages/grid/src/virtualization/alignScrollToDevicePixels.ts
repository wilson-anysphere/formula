export function alignScrollToDevicePixels(
  pos: { x: number; y: number },
  maxScroll: { maxScrollX: number; maxScrollY: number },
  devicePixelRatio: number
): { x: number; y: number } {
  const dpr = Number.isFinite(devicePixelRatio) && devicePixelRatio > 0 ? devicePixelRatio : 1;
  const step = 1 / dpr;

  const maxAlignedX = Math.floor(maxScroll.maxScrollX / step) * step;
  const maxAlignedY = Math.floor(maxScroll.maxScrollY / step) * step;

  const x = Math.min(maxAlignedX, Math.max(0, Math.round(pos.x / step) * step));
  const y = Math.min(maxAlignedY, Math.max(0, Math.round(pos.y / step) * step));
  return { x, y };
}

