export const EMU_PER_INCH = 914_400;
export const PX_PER_INCH = 96;
export const EMU_PER_PX = EMU_PER_INCH / PX_PER_INCH;

/**
 * Convert DrawingML's EMU (English Metric Unit) coordinates into CSS pixels.
 *
 * Excel / OOXML defines:
 * - 1 inch == 914400 EMU
 * - 1 CSS inch == 96 CSS pixels (per CSS spec)
 *
 * @param {number} emu
 * @returns {number}
 */
export function emuToPx(emu) {
  // Prefer the reduced ratio `EMU_PER_PX` to avoid float drift when round-tripping
  // between pixels and EMU. (E.g. `(76200 / 914400) * 96` yields `7.999999…`.)
  return emu / EMU_PER_PX;
}

/**
 * Convert CSS pixels into DrawingML EMU (English Metric Units).
 *
 * @param {number} px
 * @returns {number}
 */
export function pxToEmu(px) {
  // Prefer the reduced ratio `EMU_PER_PX` to avoid float drift for common integer
  // pixel sizes.
  return px * EMU_PER_PX;
}
