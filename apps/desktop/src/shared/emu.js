export const EMU_PER_INCH = 914_400;
export const PX_PER_INCH = 96;

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
  return (emu / EMU_PER_INCH) * PX_PER_INCH;
}

/**
 * Convert CSS pixels into DrawingML EMU (English Metric Units).
 *
 * @param {number} px
 * @returns {number}
 */
export function pxToEmu(px) {
  return (px / PX_PER_INCH) * EMU_PER_INCH;
}

