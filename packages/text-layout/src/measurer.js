import { fontKey, toCanvasFontString } from "./font.js";

/**
 * @typedef {import("./font.js").FontSpec} FontSpec
 */

/**
 * @typedef {Object} TextMeasurement
 * @property {number} width
 * @property {number} ascent
 * @property {number} descent
 */

/**
 * @typedef {Object} TextMeasurer
 * @property {(text: string, font: FontSpec) => TextMeasurement} measure
 */

export class CanvasTextMeasurer {
  /**
   * @param {CanvasRenderingContext2D} ctx
   */
  constructor(ctx) {
    this.ctx = ctx;
    this.currentFontKey = null;
  }

  /**
   * @param {string} text
   * @param {FontSpec} font
   * @returns {TextMeasurement}
   */
  measure(text, font) {
    const key = fontKey(font);
    if (this.currentFontKey !== key) {
      this.ctx.font = toCanvasFontString(font);
      this.currentFontKey = key;
    }

    const metrics = this.ctx.measureText(text);
    const ascent = metrics.actualBoundingBoxAscent ?? font.sizePx * 0.8;
    const descent = metrics.actualBoundingBoxDescent ?? font.sizePx * 0.2;
    return { width: metrics.width, ascent, descent };
  }
}

/**
 * Create a measurer backed by an internal (offscreen) canvas.
 *
 * This function is only valid in environments that support Canvas (browser / Tauri webview).
 *
 * @returns {CanvasTextMeasurer}
 */
export function createCanvasTextMeasurer() {
  /** @type {CanvasRenderingContext2D | null} */
  let ctx = null;

  if (typeof OffscreenCanvas !== "undefined") {
    const canvas = new OffscreenCanvas(1, 1);
    ctx = canvas.getContext("2d");
  } else if (typeof document !== "undefined") {
    const canvas = document.createElement("canvas");
    ctx = canvas.getContext("2d");
  }

  if (!ctx) {
    throw new Error("Canvas not available: cannot create a CanvasTextMeasurer in this environment.");
  }

  return new CanvasTextMeasurer(ctx);
}

