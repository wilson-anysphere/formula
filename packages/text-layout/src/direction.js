const RTL_SCRIPT_RE =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Adlam}]/u;
const STRONG_LTR_RE = /[\p{Alphabetic}\p{Number}]/u;

/**
 * @param {string} text
 * @returns {"ltr" | "rtl"}
 */
export function detectBaseDirection(text) {
  for (const ch of text) {
    if (RTL_SCRIPT_RE.test(ch)) return "rtl";
    if (STRONG_LTR_RE.test(ch)) return "ltr";
  }
  return "ltr";
}

/**
 * @param {"left" | "right" | "center" | "start" | "end"} align
 * @param {"ltr" | "rtl"} direction
 * @returns {"left" | "right" | "center"}
 */
export function resolveAlign(align, direction) {
  if (align === "center") return "center";
  if (align === "left" || align === "right") return align;
  if (align === "start") return direction === "rtl" ? "right" : "left";
  return direction === "rtl" ? "left" : "right";
}

