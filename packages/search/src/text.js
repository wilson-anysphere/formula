function isTypedValue(value) {
  return value != null && typeof value === "object" && typeof value.t === "string";
}

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function parseImageValue(value) {
  if (!isPlainObject(value)) return null;
  /** @type {any} */
  const obj = value;

  /** @type {any} */
  let payload = null;
  // formula-model envelope: `{ type: "image", value: {...} }`.
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
    // Direct payload shape.
    payload = obj;
  }

  if (!payload) return null;

  const imageIdRaw = payload.imageId ?? payload.image_id ?? payload.id;
  if (typeof imageIdRaw !== "string") return null;
  const imageId = imageIdRaw.trim();
  if (imageId === "") return null;

  const altTextRaw = payload.altText ?? payload.alt_text ?? payload.alt;
  const altText = typeof altTextRaw === "string" ? altTextRaw.trim() : "";

  return { imageId, altText: altText === "" ? null : altText };
}

function typedValueToString(value) {
  // The formula engine uses a stable JSON encoding (see tools/excel-oracle/value-encoding.md).
  switch (value.t) {
    case "blank":
      return "";
    case "n":
      return value.v == null ? "" : String(value.v);
    case "s":
      return value.v == null ? "" : String(value.v);
    case "b":
      // Excel displays booleans as TRUE/FALSE.
      return value.v ? "TRUE" : "FALSE";
    case "e":
      return value.v == null ? "" : String(value.v);
    case "arr":
      // Dynamic arrays/spills are represented as a matrix. There isn't a single
      // Excel "cell string" to search, but we still want stable behavior.
      // Keep this cheap (avoid deep formatting) while ensuring we don't return
      // "[object Object]" which breaks search UX.
      try {
        return JSON.stringify(value);
      } catch {
        return String(value);
      }
    default:
      // Defensive: if a newer engine introduces additional typed values, fall back
      // to stable JSON rather than leaking `[object Object]` into search results.
      try {
        const json = JSON.stringify(value);
        return typeof json === "string" ? json : String(value);
      } catch {
        return String(value);
      }
  }
}

export function formatCellValue(value) {
  if (value == null) return "";
  if (isTypedValue(value)) return typedValueToString(value);
  // Some workbook adapters (desktop DocumentController) store rich cell values (rich text,
  // in-cell images) as objects. Prefer a stable text representation over `[object Object]`.
  if (value && typeof value === "object") {
    const text = /** @type {any} */ (value)?.text;
    if (typeof text === "string") return text;

    const image = parseImageValue(value);
    if (image) return image.altText ?? "[Image]";

    // Fall back to a stable JSON representation when possible.
    try {
      const json = JSON.stringify(value);
      // Avoid emitting useless empty object text into the search index.
      if (json === "{}") return "";
      return typeof json === "string" ? json : String(value);
    } catch {
      return String(value);
    }
  }
  return String(value);
}

export function getValueText(cell, valueMode = "display") {
  if (!cell) return "";
  if (valueMode === "raw") return formatCellValue(cell.value);
  // Some adapters provide a formatted `display` field. Keep this resilient if
  // the display value is non-scalar (avoid `[object Object]`).
  if (cell.display != null) return formatCellValue(cell.display);
  return formatCellValue(cell.value);
}

export function getCellText(cell, { lookIn = "values", valueMode = "display" } = {}) {
  if (!cell) return "";

  if (lookIn === "formulas") {
    if (cell.formula != null && cell.formula !== "") return String(cell.formula);
    // Constants in "Look in: formulas" are treated as raw user input.
    return formatCellValue(cell.value);
  }

  // values
  return getValueText(cell, valueMode);
}
