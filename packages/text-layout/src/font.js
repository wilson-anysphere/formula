/**
 * @typedef {Object} FontSpec
 * @property {string} family
 * @property {number} sizePx
 * @property {string | number} [weight]
 * @property {string} [style]
 */

/**
 * @param {FontSpec} font
 * @returns {Required<Pick<FontSpec, "family" | "sizePx">> & { weight: string | number, style: string }}
 */
export function normalizeFont(font) {
  return {
    family: font.family,
    sizePx: font.sizePx,
    weight: font.weight ?? 400,
    style: font.style ?? "normal",
  };
}

/**
 * @param {FontSpec} font
 * @returns {string}
 */
export function fontKey(font) {
  const f = normalizeFont(font);
  return `${f.style}|${f.weight}|${f.sizePx}|${f.family}`;
}

/**
 * @param {FontSpec} font
 * @returns {string}
 */
export function toCanvasFontString(font) {
  const f = normalizeFont(font);
  return `${f.style} ${f.weight} ${f.sizePx}px ${f.family}`;
}

