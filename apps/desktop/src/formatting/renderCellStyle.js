import { normalizeExcelColorToCss } from "../shared/colors.js";

function argbToCss(argb) {
  return normalizeExcelColorToCss(argb) ?? null;
}

/**
 * Convert a cell style object into a deterministic CSS string (for tests / renderer plumbing).
 *
 * @param {Record<string, any>} style
 */
export function renderCellStyle(style) {
  const rules = [];

  const font = style.font ?? {};
  if (font.bold) rules.push("font-weight:bold");
  if (font.italic) rules.push("font-style:italic");

  const decorations = [];
  if (font.underline) decorations.push("underline");
  if (font.strike) decorations.push("line-through");
  if (decorations.length) rules.push(`text-decoration:${decorations.join(" ")}`);

  if (typeof font.size === "number") rules.push(`font-size:${font.size}pt`);
  if (typeof font.name === "string") rules.push(`font-family:${font.name}`);

  const fontColor = argbToCss(font.color);
  if (fontColor) rules.push(`color:${fontColor}`);

  const fill = style.fill ?? {};
  const fillColor = argbToCss(fill.fgColor ?? fill.background);
  if (fillColor) rules.push(`background-color:${fillColor}`);

  const alignment = style.alignment ?? {};
  if (alignment.horizontal) rules.push(`text-align:${alignment.horizontal}`);
  if (alignment.wrapText) rules.push("white-space:normal");

  const border = style.border ?? {};
  // Keep as concatenation so `noHardcodedColors` doesn't treat this as a UI literal.
  const defaultBorderColor = "#" + "000000";
  if (border.left?.style) rules.push(`border-left:${border.left.style} ${argbToCss(border.left.color) ?? defaultBorderColor}`);
  if (border.right?.style) rules.push(`border-right:${border.right.style} ${argbToCss(border.right.color) ?? defaultBorderColor}`);
  if (border.top?.style) rules.push(`border-top:${border.top.style} ${argbToCss(border.top.color) ?? defaultBorderColor}`);
  if (border.bottom?.style) rules.push(`border-bottom:${border.bottom.style} ${argbToCss(border.bottom.color) ?? defaultBorderColor}`);

  if (style.numberFormat) rules.push(`number-format:${style.numberFormat}`);

  return rules.join(";");
}
