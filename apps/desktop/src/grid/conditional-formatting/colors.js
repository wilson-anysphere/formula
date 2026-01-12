import { normalizeExcelColorToCss } from "../../shared/colors.js";

/**
 * @param {unknown} argb
 * @returns {string | undefined}
 */
export function argbToCss(argb) {
  return normalizeExcelColorToCss(argb);
}
