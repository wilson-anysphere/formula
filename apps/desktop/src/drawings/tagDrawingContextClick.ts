export type DrawingHitResult = { id: number };

export type DrawingContextClickHitTester = {
  hitTestDrawingAtClientPoint(clientX: number, clientY: number): DrawingHitResult | null;
};

export function resolveIsMacPlatform(): boolean {
  try {
    const platform = typeof navigator !== "undefined" ? navigator.platform : "";
    return /Mac|iPhone|iPad|iPod/.test(platform);
  } catch {
    return false;
  }
}

/**
 * Tags a pointer event that context-clicked a drawing so CanvasGrid-based selection handlers can
 * ignore it (preventing the active cell from moving underneath the drawing).
 *
 * Returns true when the event was tagged.
 */
export function tagDrawingContextClickPointerDown(
  event: PointerEvent,
  app: DrawingContextClickHitTester,
  opts: { isMacPlatform?: boolean; requireCanvasTarget?: boolean } = {},
): boolean {
  const isMacPlatform = opts.isMacPlatform ?? resolveIsMacPlatform();
  const requireCanvasTarget = opts.requireCanvasTarget !== false;

  if (event.pointerType !== "mouse") return false;
  const isContextClick = event.button === 2 || (isMacPlatform && event.button === 0 && event.ctrlKey && !event.metaKey);
  if (!isContextClick) return false;

  if (requireCanvasTarget) {
    const target = event.target as HTMLElement | null;
    if (!(target instanceof HTMLCanvasElement)) return false;
  }

  const hit = app.hitTestDrawingAtClientPoint(event.clientX, event.clientY);
  if (!hit) return false;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (event as any).__formulaDrawingContextClick = true;
  return true;
}

