import { createCanvasTextMeasurer, TextLayoutEngine } from "@formula/text-layout";

let sharedEngine = null;

/**
 * @returns {TextLayoutEngine}
 */
export function getSharedTextLayoutEngine() {
  if (sharedEngine) return sharedEngine;
  sharedEngine = new TextLayoutEngine(createCanvasTextMeasurer());
  return sharedEngine;
}

