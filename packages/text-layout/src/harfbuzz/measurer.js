import GraphemeSplitter from "grapheme-splitter";

import { LRUCache } from "../lru.js";

/**
 * @typedef {import("../measurer.js").TextMeasurement} TextMeasurement
 * @typedef {import("../font.js").FontSpec} FontSpec
 * @typedef {import("./font-manager.js").HarfBuzzFontManager} HarfBuzzFontManager
 * @typedef {import("./font-manager.js").HarfBuzzFontFace} HarfBuzzFontFace
 */

const GRAPHEME_SPLITTER = new GraphemeSplitter();

export class HarfBuzzTextMeasurer {
  /**
   * @param {HarfBuzzFontManager} fontManager
   * @param {{ maxShapeCacheEntries?: number }} [opts]
   */
  constructor(fontManager, opts = {}) {
    this.fontManager = fontManager;

    // Cache the sum of advances for (faceKey,text) in design units. `TextLayoutEngine` already
    // caches `(fontKey,text)` at a higher level, but this keeps `HarfBuzzTextMeasurer` usable on its own.
    this.shapeCache = new LRUCache(opts.maxShapeCacheEntries ?? 10_000);
  }

  get cacheKey() {
    // Incorporate font-manager version so `TextLayoutEngine` caches are invalidated when fonts or
    // fallback configuration changes.
    return `harfbuzz:v${this.fontManager.version}`;
  }

  /**
   * @param {string} text
   * @param {FontSpec} font
   * @returns {TextMeasurement}
   */
  measure(text, font) {
    if (!text) return { width: 0, ascent: 0, descent: 0 };

    const sizePx = font.sizePx;
    const primary = this.fontManager.getFace(font);
    const fallbacks = this.fontManager.getFallbackFaces(font);

    const segments = primary.supportsText(text)
      ? [{ face: primary, text }]
      : this.#segmentByFont(text, primary, fallbacks);

    let width = 0;
    let ascent = 0;
    let descent = 0;

    for (const seg of segments) {
      const advanceUnits = this.#shapeAdvanceUnits(seg.face, seg.text);
      width += (advanceUnits * sizePx) / seg.face.upem;
      ascent = Math.max(ascent, seg.face.ascentRatio * sizePx);
      descent = Math.max(descent, seg.face.descentRatio * sizePx);
    }

    return { width, ascent, descent };
  }

  /**
   * @param {HarfBuzzFontFace} face
   * @param {string} text
   * @returns {number} Sum of glyph advances in font design units.
   */
  #shapeAdvanceUnits(face, text) {
    const cacheKey = `${this.fontManager.version}|${face.key}\n${text}`;
    const cached = this.shapeCache.get(cacheKey);
    if (cached !== undefined) return cached;

    const hb = this.fontManager.hb;
    const buffer = hb.createBuffer();
    buffer.addText(text);
    buffer.guessSegmentProperties();
    hb.shape(face.hbFont, buffer);

    let sum = 0;
    for (const g of buffer.json()) sum += g.ax;
    buffer.destroy();

    this.shapeCache.set(cacheKey, sum);
    return sum;
  }

  /**
   * Segment a string into grapheme clusters and assign each cluster to a font that can render it.
   *
   * @param {string} text
   * @param {HarfBuzzFontFace} primary
   * @param {HarfBuzzFontFace[]} fallbacks
   * @returns {Array<{ face: HarfBuzzFontFace, text: string }>}
   */
  #segmentByFont(text, primary, fallbacks) {
    const clusters = GRAPHEME_SPLITTER.splitGraphemes(text);

    /** @type {Array<{ face: HarfBuzzFontFace, text: string }>} */
    const segments = [];

    /** @type {HarfBuzzFontFace | null} */
    let currentFace = null;
    let currentText = "";

    const flush = () => {
      if (currentFace && currentText) segments.push({ face: currentFace, text: currentText });
      currentText = "";
    };

    for (const cluster of clusters) {
      /** @type {HarfBuzzFontFace} */
      let face = primary;
      if (!primary.supportsText(cluster)) {
        face = fallbacks.find((f) => f.supportsText(cluster)) ?? primary;
      }

      if (!currentFace) {
        currentFace = face;
        currentText = cluster;
        continue;
      }

      if (face.key === currentFace.key) {
        currentText += cluster;
      } else {
        flush();
        currentFace = face;
        currentText = cluster;
      }
    }

    flush();
    return segments;
  }
}
