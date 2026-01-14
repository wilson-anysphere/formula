import { cellToA1, rectToA1 } from "./rect.js";

const DEFAULT_MAX_COLUMNS_FOR_SCHEMA = 20;
const DEFAULT_MAX_COLUMNS_FOR_ROWS = 20;
const MAX_FORMULA_SAMPLES = 12;

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Convert an arbitrary value into a safe JSON-serializable structure without invoking
 * user-controlled stringification hooks like `toJSON()` / `toString()`.
 *
 * This is a defense-in-depth helper: workbook cells can contain rich objects (including
 * class instances) where `toJSON`/`toString` may be overridden to leak sensitive strings or
 * throw. We prefer a stable placeholder over calling those hooks.
 *
 * @param {unknown} value
 * @param {WeakSet<object>} stack
 * @returns {unknown}
 */
function safeJsonValue(value, stack) {
  if (value === null || value === undefined) return null;
  const t = typeof value;
  if (t === "string" || t === "number" || t === "boolean") return value;
  if (t === "bigint") return value.toString();
  if (t === "symbol") return "Symbol";
  if (t === "function") return "Function";

  if (t !== "object") return null;

  if (value instanceof Date) {
    // Avoid calling per-instance overrides (e.g. `date.toISOString = () => "secret"`).
    let time = NaN;
    try {
      time = Date.prototype.getTime.call(value);
    } catch {
      time = NaN;
    }
    if (!Number.isFinite(time)) return "";
    try {
      return Date.prototype.toISOString.call(value);
    } catch {
      return "";
    }
  }

  if (value instanceof Error) {
    const message = typeof value.message === "string" ? value.message.trim() : "";
    const name = typeof value.name === "string" && value.name.trim() ? value.name.trim() : "Error";
    return message ? `${name}: ${message}` : name;
  }

  // Avoid scanning large binary blobs.
  if (typeof ArrayBuffer !== "undefined") {
    if (value instanceof ArrayBuffer) return "[Binary]";
    if (typeof ArrayBuffer.isView === "function" && ArrayBuffer.isView(value)) return "[Binary]";
  }

  if (stack.has(value)) return "[Circular]";
  stack.add(value);
  try {
    if (Array.isArray(value)) return value.map((v) => safeJsonValue(v, stack));
    if (value instanceof Map) {
      return Array.from(value.entries()).map(([k, v]) => [safeJsonValue(k, stack), safeJsonValue(v, stack)]);
    }
    if (value instanceof Set) {
      return Array.from(value.values()).map((v) => safeJsonValue(v, stack));
    }

    // Plain object / class instance: serialize enumerable fields, dropping callable hooks.
    /** @type {Record<string, unknown>} */
    const out = {};
    const obj = /** @type {any} */ (value);
    for (const key of Object.keys(obj).sort()) {
      const v = obj[key];
      // Drop stringification hooks so downstream JSON.stringify can't call them.
      if ((key === "toString" || key === "valueOf" || key === "toJSON") && typeof v === "function") continue;
      if (typeof v === "function") continue;
      out[key] = safeJsonValue(v, stack);
    }
    return out;
  } finally {
    stack.delete(value);
  }
}

/**
 * Safe JSON stringification for cell values.
 *
 * Unlike `JSON.stringify(value)` this avoids invoking `toJSON()`, which is user-controlled.
 *
 * @param {unknown} value
 * @returns {string | null}
 */
function safeJsonStringify(value) {
  try {
    return JSON.stringify(safeJsonValue(value, new WeakSet()));
  } catch {
    return null;
  }
}

function parseImageValue(value) {
  if (!isPlainObject(value)) return null;
  const obj = /** @type {any} */ (value);

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
  let altText = null;
  if (typeof altTextRaw === "string") {
    const trimmed = altTextRaw.trim();
    if (trimmed !== "") altText = trimmed;
  }

  return { imageId, altText };
}

function formatScalar(value) {
  if (value == null) return "";
  if (value instanceof Date) {
    // Avoid calling per-instance overrides (e.g. `date.toISOString = () => "secret"`).
    let time = NaN;
    try {
      time = Date.prototype.getTime.call(value);
    } catch {
      time = NaN;
    }
    if (!Number.isFinite(time)) return "";
    // Use ISO format for stability and compactness (better embeddings than
    // Date#toString locale/timezone output).
    try {
      return Date.prototype.toISOString.call(value);
    } catch {
      return "";
    }
  }
  if (typeof value === "bigint") {
    // BigInts can be extremely long; reuse the string truncation path.
    return formatScalar(String(value));
  }
  if (value instanceof Error) {
    const message = typeof value.message === "string" ? value.message.trim() : "";
    const name = typeof value.name === "string" && value.name.trim() ? value.name.trim() : "Error";
    // Re-run through the string path so we inherit whitespace normalization,
    // pipe escaping, and truncation.
    return formatScalar(message ? `${name}: ${message}` : name);
  }
  if (typeof value === "object") {
    // Some backends surface rich cell values (e.g. structured types / rich text).
    // Prefer a stable, compact string representation over "[object Object]".
    const text = /** @type {any} */ (value)?.text;
    if (typeof text === "string") return formatScalar(text);
    const image = parseImageValue(value);
    if (image) return formatScalar(image.altText ?? "[Image]");
    const json = safeJsonStringify(value);
    if (typeof json === "string") {
      // Empty objects are rarely useful in cell context; treat as blank.
      if (json === "{}") return "";
      return formatScalar(json);
    }
    // Some objects (circular refs, BigInt, etc) aren't JSON stringifiable.
    // Fall back to a stable placeholder rather than "[object Object]".
    return "Object";
  }
  if (typeof value === "string") {
    const trimmed = value.replace(/\s+/g, " ").trim();
    // Our output format uses `|` as a column separator; replace literal pipes in
    // cell text so the rendered rows remain unambiguous in LLM context.
    const cleaned = trimmed.replace(/\|/g, "¦");
    if (cleaned.length > 60) return `${cleaned.slice(0, 57)}...`;
    return cleaned;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return String(value);
    // Keep stable precision without noise.
    return Number.isInteger(value) ? String(value) : value.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "symbol") return "Symbol";
  if (typeof value === "function") return "Function";
  return String(value);
}

/**
 * Headers are rendered into `Header=value` pairs. Escape `=` so the output stays
 * parseable when headers contain equals signs.
 *
 * @param {string} header
 */
function sanitizeHeaderLabel(header) {
  return String(header).replace(/=/g, "≡").trim();
}

/**
 * Ensure header labels are unique so sample rows can be parsed as key/value pairs
 * without ambiguity (duplicate header names are common in messy spreadsheets).
 *
 * @param {string[]} headers
 */
function dedupeHeaders(headers) {
  /** @type {Map<string, number>} */
  const seen = new Map();
  return headers.map((h) => {
    const base = String(h);
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    if (count === 0) return base;
    return `${base}_${count + 1}`;
  });
}

/**
 * @param {number} total
 * @param {number} shown
 */
function formatExtraColumns(total, shown) {
  const extra = total - shown;
  if (extra <= 0) return null;
  return `… (+${extra} more columns)`;
}

/**
 * @param {any[][]} cells
 * @param {number} col
 * @param {number} headerRow
 */
function inferColumnType(cells, col, headerRow) {
  const sample = [];
  for (let r = headerRow + 1; r < Math.min(cells.length, headerRow + 21); r += 1) {
    const cell = cells[r]?.[col];
    const v = cell?.v;
    if (v == null || v === "") continue;
    sample.push(v);
  }
  if (sample.length === 0) return "empty";
  let hasNumber = false;
  let hasString = false;
  let hasBool = false;
  for (const v of sample) {
    if (typeof v === "number") hasNumber = true;
    else if (typeof v === "boolean") hasBool = true;
    else hasString = true;
  }
  const kindCount = Number(hasNumber) + Number(hasString) + Number(hasBool);
  if (kindCount > 1) return "mixed";
  if (hasNumber) return "number";
  if (hasBool) return "boolean";
  return "string";
}

/**
 * @param {any[][]} cells
 */
function inferHeaderRow(cells) {
  if (!cells.length) return null;
  const maxRowsToCheck = Math.min(cells.length, 5);
  let maxColCount = 0;
  for (let r = 0; r < maxRowsToCheck; r += 1) {
    const row = cells[r] || [];
    if (Array.isArray(row) && row.length > maxColCount) maxColCount = row.length;
  }

  /**
   * @param {number} rowIndex
   */
  function rowStats(rowIndex) {
    const row = cells[rowIndex] || [];
    let nonEmpty = 0;
    let stringish = 0;
    let firstString = null;
    for (const cell of row) {
      const v = cell?.v;
      if (v == null) continue;

      // For header detection we want to treat rich values (e.g. DocumentController rich text + in-cell
      // images) as text, while still treating numbers/booleans as non-string values so purely numeric
      // rows are not misclassified as headers.
      let text = null;
      if (typeof v === "string") {
        text = v;
      } else if (isPlainObject(v)) {
        const maybeText = /** @type {any} */ (v)?.text;
        if (typeof maybeText === "string") {
          text = maybeText;
        } else {
          const image = parseImageValue(v);
          if (image) text = image.altText ?? "[Image]";
        }
      }

      if (text != null) {
        const trimmedText = text.trim();
        if (!trimmedText) continue;
        nonEmpty += 1;
        stringish += 1;
        if (firstString == null) firstString = trimmedText;
        continue;
      }

      // Avoid calling `String(v)` on arbitrary objects; custom `toString()` implementations can
      // throw or leak sensitive strings. For header-row heuristics we only need to know whether
      // the cell is non-empty, not its rendered text.
      const scalarType = typeof v;
      if (scalarType === "string") {
        // Should be handled by the `text != null` branch above, but keep this defensive.
        const trimmed = v.trim();
        if (!trimmed) continue;
        nonEmpty += 1;
        continue;
      }
      if (scalarType === "number") {
        if (!Number.isFinite(v)) continue;
        nonEmpty += 1;
        continue;
      }
      if (scalarType === "boolean" || scalarType === "bigint" || scalarType === "symbol" || scalarType === "function") {
        nonEmpty += 1;
        continue;
      }
      if (scalarType === "object") {
        // Any remaining object-like value counts as non-empty.
        nonEmpty += 1;
        continue;
      }
    }
    return { row, nonEmpty, stringish, firstString };
  }

  /**
   * @param {{ nonEmpty: number, stringish: number }} stats
   */
  function isHeaderCandidate(stats) {
    if (stats.nonEmpty === 0) return false;
    return stats.stringish / stats.nonEmpty >= 0.6;
  }

  const row0 = rowStats(0);
  const row0IsCandidate = isHeaderCandidate(row0);
  if (row0IsCandidate) {
    // Special-case "title rows": a single long-ish multi-word string in the first
    // row of an otherwise multi-column range often indicates a caption above the
    // actual header row (e.g. "Revenue Summary").
    //
    // Be conservative: it's easy for a real header to have a single multi-word
    // label in the first column with the remaining headers blank (e.g.
    // "Customer Name"). Prefer false negatives (treat as a header) over false
    // positives (treating data as headers).
    const titleKeywordsRe = /\b(summary|report|overview|dashboard|analysis|results|totals)\b/i;
    const firstString = row0.firstString;
    const keywordTitle = typeof firstString === "string" && titleKeywordsRe.test(firstString);
    const punctTitle = typeof firstString === "string" && /[:–—]/.test(firstString);
    const hasSpaces = typeof firstString === "string" && /\s/.test(firstString);

    // If the first row is a single keyword-y label like "Summary" / "Report",
    // only treat it as a title row if a later header candidate looks multi-column.
    let laterMultiColHeader = false;
    if (keywordTitle) {
      for (let r = 1; r < maxRowsToCheck; r += 1) {
        const stats = rowStats(r);
        if (!isHeaderCandidate(stats)) continue;
        if (stats.nonEmpty >= 2) {
          laterMultiColHeader = true;
          break;
        }
      }
    }

    const titleLike =
      row0.nonEmpty === 1 &&
      maxColCount > 1 &&
      typeof firstString === "string" &&
      ((hasSpaces &&
        firstString.length >= 12 &&
        (keywordTitle || firstString.length >= 24 || punctTitle)) ||
        (keywordTitle && laterMultiColHeader));
    if (!titleLike) return 0;
  }

  let bestRow = null;
  let bestStringish = 0;
  let bestNonEmpty = 0;
  for (let r = row0IsCandidate ? 1 : 0; r < maxRowsToCheck; r += 1) {
    const stats = rowStats(r);
    if (!isHeaderCandidate(stats)) continue;
    if (stats.stringish > bestStringish || (stats.stringish === bestStringish && stats.nonEmpty > bestNonEmpty)) {
      bestRow = r;
      bestStringish = stats.stringish;
      bestNonEmpty = stats.nonEmpty;
    }
  }
  return bestRow;
}

function countFormulas(cells) {
  let count = 0;
  for (const row of cells) {
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      if (!cell || cell.f == null) continue;
      const f = cell.f;
      if (typeof f === "string") {
        if (f.trim() !== "") count += 1;
        continue;
      }
      if (typeof f === "number") {
        if (Number.isFinite(f)) count += 1;
        continue;
      }
      if (typeof f === "boolean" || typeof f === "bigint" || typeof f === "symbol" || typeof f === "function") {
        count += 1;
        continue;
      }
      if (typeof f === "object") {
        count += 1;
      }
    }
  }
  return count;
}

/**
 * Convert a workbook chunk into a compact, schema-first representation suitable
 * for RAG embedding and LLM context injection.
 *
 * @param {import('./workbookTypes').WorkbookChunk} chunk
 * @param {{ sampleRows?: number, maxColumnsForSchema?: number, maxColumnsForRows?: number }} [opts]
 */
export function chunkToText(chunk, opts) {
  const sampleRows = opts?.sampleRows ?? 5;
  const maxColumnsForSchema = opts?.maxColumnsForSchema ?? DEFAULT_MAX_COLUMNS_FOR_SCHEMA;
  const maxColumnsForRows = opts?.maxColumnsForRows ?? DEFAULT_MAX_COLUMNS_FOR_ROWS;
  const rectA1 = rectToA1(chunk.rect);
  const cells = chunk.cells || [];
  const headerRow = inferHeaderRow(cells);
  let sampledColCount = 0;
  for (const row of cells) {
    if (Array.isArray(row) && row.length > sampledColCount) sampledColCount = row.length;
  }
  const sampledRowCount = cells.length;
  const fullColCount = chunk.rect.c1 - chunk.rect.c0 + 1;
  const fullRowCount = chunk.rect.r1 - chunk.rect.r0 + 1;
  const formulaCount = countFormulas(cells);

  const kind =
    typeof chunk?.kind === "string" && chunk.kind.trim() ? chunk.kind.trim().toUpperCase() : "CHUNK";
  const title = typeof chunk?.title === "string" ? chunk.title : "";
  // Sheet names are expected to be strings; avoid coercing arbitrary objects via `String(...)`.
  const sheetName = typeof chunk?.sheetName === "string" ? chunk.sheetName : "";

  const lines = [];
  lines.push(
    `${kind}: ${title} (sheet="${sheetName}", range="${rectA1}", size=${fullRowCount}x${fullColCount}, formulas≈${formulaCount})`
  );

  if (sampledRowCount < fullRowCount || sampledColCount < fullColCount) {
    lines.push(
      `NOTE: embedding uses a ${sampledRowCount}x${sampledColCount} cell sample (full range is ${fullRowCount}x${fullColCount}).`
    );
  }

  const schemaColCount = Math.max(0, Math.min(sampledColCount, maxColumnsForSchema));
  const rowColCount = Math.max(0, Math.min(sampledColCount, maxColumnsForRows));
  const headerNames =
    headerRow != null
      ? dedupeHeaders(
          Array.from({ length: Math.max(schemaColCount, rowColCount) }, (_, c) => {
          const h = formatScalar(cells[headerRow]?.[c]?.v);
          return sanitizeHeaderLabel(h || `Column${c + 1}`);
        })
        )
      : null;

  if (sampledColCount > 0) {
    if (headerRow != null) {
      const headers = [];
      const types = [];
      for (let c = 0; c < schemaColCount; c += 1) {
        headers.push(headerNames?.[c] ?? `Column${c + 1}`);
        types.push(inferColumnType(cells, c, headerRow));
      }
      const extra = formatExtraColumns(fullColCount, schemaColCount);
      if (extra) headers.push(extra);
      lines.push(`COLUMNS: ${headers.map((h, i) => (types[i] ? `${h} (${types[i]})` : h)).join(" | ")}`);
    } else {
      const parts = [];
      for (let c = 0; c < schemaColCount; c += 1) {
        const type = inferColumnType(cells, c, -1);
        parts.push(`Column${c + 1} (${type})`);
      }
      const extra = formatExtraColumns(fullColCount, schemaColCount);
      if (extra) parts.push(extra);
      lines.push(`COLUMNS: ${parts.join(" | ")}`);
    }
  }

  if (headerRow != null && headerRow > 0 && chunk.kind !== "formulaRegion") {
    const preRows = [];
    const maxPreRowsToShow = 2;
    for (let r = 0; r < headerRow && preRows.length < maxPreRowsToShow; r += 1) {
      const values = [];
      for (let c = 0; c < rowColCount; c += 1) {
        const cell = cells[r]?.[c] || {};
        if (cell.f) {
          const formula = formatScalar(cell.f);
          const value = formatScalar(cell.v);
          values.push(value ? `${formula}=${value}` : formula);
        } else {
          values.push(formatScalar(cell.v));
        }
      }
      const extra = formatExtraColumns(fullColCount, rowColCount);
      if (extra) values.push(extra);
      const compact = values.filter((v) => v !== "").join(" | ");
      if (compact) preRows.push(compact);
    }
    if (preRows.length) {
      lines.push("PRE-HEADER ROWS:");
      for (const row of preRows) lines.push(`  - ${row}`);
      const remaining = headerRow - preRows.length;
      if (remaining > 0) lines.push(`  - … (+${remaining} more pre-header rows)`);
    }
  }

  if (chunk.kind === "formulaRegion") {
    const formulas = [];
    for (let r = 0; r < cells.length && formulas.length < MAX_FORMULA_SAMPLES; r += 1) {
      for (let c = 0; c < (cells[r]?.length ?? 0) && formulas.length < MAX_FORMULA_SAMPLES; c += 1) {
        const cell = cells[r][c] || {};
        const f = cell.f;
        if (!f) continue;
        const addr = cellToA1(chunk.rect.r0 + r, chunk.rect.c0 + c);
        const value = formatScalar(cell.v);
        const formula = formatScalar(f);
        formulas.push(value ? `${addr}:${formula}=${value}` : `${addr}:${formula}`);
      }
    }
    const extraFormulas = formulaCount - formulas.length;
    if (extraFormulas > 0) formulas.push(`… (+${extraFormulas} more formulas)`);
    if (formulas.length) lines.push(`FORMULAS: ${formulas.join(" | ")}`);
  } else {
    const startRow = headerRow != null ? headerRow + 1 : 0;
    const rows = [];
    for (let r = startRow; r < Math.min(sampledRowCount, startRow + sampleRows); r += 1) {
      const row = [];
      for (let c = 0; c < rowColCount; c += 1) {
        const cell = cells[r]?.[c] || {};
        if (headerNames) {
          const header = headerNames[c] || `Column${c + 1}`;
          if (cell.f) {
            const formula = formatScalar(cell.f);
            const value = formatScalar(cell.v);
            row.push(value ? `${header}(${formula})=${value}` : `${header}(${formula})`);
          } else {
            row.push(`${header}=${formatScalar(cell.v)}`);
          }
        } else {
          if (cell.f) {
            const formula = formatScalar(cell.f);
            const value = formatScalar(cell.v);
            row.push(value ? `${formula}=${value}` : formula);
          } else {
            row.push(formatScalar(cell.v));
          }
        }
      }
      const extra = formatExtraColumns(fullColCount, rowColCount);
      if (extra) row.push(extra);
      rows.push(row.join(" | "));
    }
    const extraRows = fullRowCount - startRow - rows.length;
    if (rows.length > 0 && extraRows > 0) rows.push(`… (+${extraRows} more rows)`);
    if (rows.length) {
      lines.push("SAMPLE ROWS:");
      for (const row of rows) lines.push(`  - ${row}`);
    }
  }

  return lines.join("\n");
}
