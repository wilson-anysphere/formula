import { CanvasTextMeasurer, createCanvasTextMeasurer, TextLayoutEngine } from "@formula/text-layout";

let sharedEngine = null;

/**
 * Create (and cache) a single `TextLayoutEngine` instance for the lifetime of the app.
 *
 * In browser/Tauri environments we prefer an internal offscreen canvas measurer so layout
 * does not mutate the render context state. In non-browser test environments where Canvas
 * isn't available we fall back to the provided context if it implements `measureText`.
 *
 * @param {any} [fallbackCtx]
 * @returns {TextLayoutEngine}
 */
export function getSharedTextLayoutEngine(fallbackCtx) {
  if (sharedEngine) return sharedEngine;
  let measurer;
  try {
    measurer = createCanvasTextMeasurer();
  } catch (error) {
    if (fallbackCtx && typeof fallbackCtx.measureText === "function") {
      measurer = new CanvasTextMeasurer(fallbackCtx);
    } else {
      throw error;
    }
  }

  sharedEngine = new TextLayoutEngine(measurer, {
    maxMeasureCacheEntries: 50_000,
    maxLayoutCacheEntries: 10_000,
  });
  return sharedEngine;
}
