/**
 * Lightweight helpers for the desktop "in-cell image" cell value payloads.
 *
 * These values originate from XLSX RichData extraction (Excel "Place in Cell" pictures
 * and IMAGE() rich-value caches) and are stored in the DocumentController as a JSON-ish
 * object so the shared grid renderer can resolve image bytes by id.
 *
 * The broader app has many places that stringify cell values for copy/export/extension APIs.
 * This helper keeps those paths from rendering as "[object Object]".
 *
 * @typedef {{ imageId: string, altText: string | null }} ParsedImageCellValue
 */

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Parse a DocumentController cell value into a normalized in-cell image payload.
 *
 * Supported shapes:
 * - `{ type: "image", value: { imageId, altText? } }` (formula-model style envelope)
 * - `{ imageId, altText? }` (direct payload / legacy variants with snake_case keys)
 *
 * @param {unknown} value
 * @returns {ParsedImageCellValue | null}
 */
export function parseImageCellValue(value) {
  if (!isPlainObject(value)) return null;
  const obj = /** @type {any} */ (value);

  let payload = null;
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
    payload = obj;
  }
  if (!payload) return null;

  const imageId = payload.imageId ?? payload.image_id ?? payload.id;
  if (typeof imageId !== "string" || imageId.trim() === "") return null;

  const altTextRaw = payload.altText ?? payload.alt_text ?? payload.alt;
  const altText = typeof altTextRaw === "string" ? altTextRaw.trim() : "";

  return { imageId: imageId.trim(), altText: altText === "" ? null : altText };
}
