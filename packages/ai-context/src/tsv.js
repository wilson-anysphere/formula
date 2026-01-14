import { isCellEmpty } from "./a1.js";
import { throwIfAborted } from "./abort.js";

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function isTypedValue(value) {
  return isPlainObject(value) && typeof /** @type {any} */ (value).t === "string";
}

/**
 * @param {unknown} value
 * @returns {{ imageId: string, altText: string | null } | null}
 */
function parseImageValue(value) {
  if (!isPlainObject(value)) return null;
  const obj = /** @type {any} */ (value);

  let payload = null;
  // DocumentController / formula-model envelope: `{ type: "image", value: {...} }`.
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
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

function typedValueToTsv(value) {
  const v = /** @type {any} */ (value);
  switch (v.t) {
    case "blank":
      return "";
    case "n":
      return v.v == null ? "" : String(v.v);
    case "s":
      return v.v == null ? "" : String(v.v);
    case "b":
      // Excel uses TRUE/FALSE for boolean display.
      return v.v ? "TRUE" : "FALSE";
    case "e":
      return v.v == null ? "" : String(v.v);
    case "arr":
      // Avoid stringifying potentially huge matrices.
      return "[array]";
    default:
      // Defensive: preserve embedded scalar payloads if present and avoid "[object Object]".
      if (Object.prototype.hasOwnProperty.call(v, "v")) return v.v == null ? "" : String(v.v);
      try {
        const json = JSON.stringify(value);
        return typeof json === "string" ? json : "";
      } catch {
        return "Object";
      }
  }
}

function formatValueForTsv(value) {
  if (isCellEmpty(value)) return "";

  if (isTypedValue(value)) return typedValueToTsv(value);

  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "bigint") return value.toString();

  if (typeof value === "object") {
    // DocumentController rich text values: `{ text, runs }`.
    const text = /** @type {any} */ (value)?.text;
    if (typeof text === "string") return text;

    const image = parseImageValue(value);
    if (image) return image.altText ?? "[Image]";

    // Treat `{}` as blank; it's a common sparse representation.
    if (value && value.constructor === Object && Object.keys(value).length === 0) return "";

    if (value instanceof Date) {
      try {
        // Avoid calling per-instance overrides (e.g. `date.toISOString = () => "secret"`).
        return Date.prototype.toISOString.call(value);
      } catch {
        // Invalid dates throw in `toISOString()`; fall back to a stable, non-throwing string form.
        try {
          return Date.prototype.toString.call(value);
        } catch {
          return "";
        }
      }
    }

    // Prefer the object's own stringification if it yields something more meaningful than
    // the default "[object Object]". This also preserves callsites that use custom `toString`
    // implementations to trigger abort signals in the middle of long TSV renders.
    let textified = "";
    try {
      textified = String(value);
    } catch {
      textified = "";
    }
    if (textified && textified !== "[object Object]") return textified;

    // Stable representation for other object-like values (avoid leaking "[object Object]").
    try {
      const json = JSON.stringify(value);
      if (json === "{}") return "";
      return typeof json === "string" ? json : textified || "Object";
    } catch {
      return textified && textified !== "[object Object]" ? textified : "Object";
    }
  }

  try {
    return String(value);
  } catch {
    return "";
  }
}

/**
 * Convert a sub-range of a sheet's value matrix to TSV.
 *
 * This is intentionally streaming-ish: it only reads up to `maxRows` rows from `values`
 * rather than allocating a full `slice2D()` copy of the entire range.
 *
 * @param {unknown[][]} values
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 * @param {{ maxRows: number, signal?: AbortSignal }} options
 */
export function valuesRangeToTsv(values, range, options) {
  const signal = options.signal;
  const shouldCheckAbort = Boolean(signal);
  const lines = [];
  const totalRows = range.endRow - range.startRow + 1;
  const limit = Math.min(totalRows, options.maxRows);

  for (let rOffset = 0; rOffset < limit; rOffset++) {
    if (shouldCheckAbort) throwIfAborted(signal);
    const row = values[range.startRow + rOffset];
    if (!Array.isArray(row)) {
      lines.push("");
      continue;
    }

    // Preserve `slice2D(...)+matrixToTsv(...)` ragged-row semantics:
    // only include columns that exist in the source row slice.
    const rowLen = row.length;
    if (rowLen <= range.startCol) {
      lines.push("");
      continue;
    }

    const sliceLen = Math.max(0, Math.min(rowLen, range.endCol + 1) - range.startCol);
    if (sliceLen === 0) {
      lines.push("");
      continue;
    }

    /** @type {string[]} */
    const cells = new Array(sliceLen);
    for (let cOffset = 0; cOffset < sliceLen; cOffset++) {
      // Avoid calling `throwIfAborted` for every cell when no signal is provided.
      // When a signal exists, check periodically to keep cancellation responsive
      // even for very wide ranges.
      if (shouldCheckAbort && (cOffset & 0x7f) === 0) throwIfAborted(signal);
      const v = row[range.startCol + cOffset];
      cells[cOffset] = formatValueForTsv(v);
    }
    lines.push(cells.join("\t"));
  }

  if (totalRows > limit) lines.push(`… (${totalRows - limit} more rows)`);
  return lines.join("\n");
}
