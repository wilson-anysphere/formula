import { loadHarfBuzz } from "./loader.js";
import { HarfBuzzFontManager } from "./font-manager.js";
import { HarfBuzzTextMeasurer } from "./measurer.js";

export { loadHarfBuzz, HarfBuzzFontManager, HarfBuzzTextMeasurer };

/**
 * @typedef {import("../font.js").FontSpec} FontSpec
 */

/**
 * Create a HarfBuzz-backed measurer with an empty `HarfBuzzFontManager`.
 *
 * The returned measurer is ready for use once fonts are loaded into `measurer.fontManager`.
 *
 * @param {{
 *   fonts?: Array<(Omit<FontSpec, "sizePx"> & { data: ArrayBuffer | ArrayBufferView })>,
 *   fallbackFamilies?: string[],
 *   maxShapeCacheEntries?: number,
 * }} [opts]
 * @returns {Promise<HarfBuzzTextMeasurer>}
 */
export async function createHarfBuzzTextMeasurer(opts = {}) {
  const hb = await loadHarfBuzz();
  const fontManager = new HarfBuzzFontManager(hb);
  if (opts.fallbackFamilies) fontManager.setFallbackFamilies(opts.fallbackFamilies);
  if (opts.fonts) {
    for (const f of opts.fonts) {
      fontManager.loadFont(f.data, { family: f.family, weight: f.weight, style: f.style });
    }
  }
  return new HarfBuzzTextMeasurer(fontManager, { maxShapeCacheEntries: opts.maxShapeCacheEntries });
}
