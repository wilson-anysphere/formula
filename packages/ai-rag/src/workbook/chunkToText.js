import { cellToA1, rectToA1 } from "./rect.js";

const DEFAULT_MAX_COLUMNS_FOR_SCHEMA = 20;
const DEFAULT_MAX_COLUMNS_FOR_ROWS = 20;
const MAX_FORMULA_SAMPLES = 12;

function formatScalar(value) {
  if (value == null) return "";
  if (typeof value === "string") {
    const trimmed = value.replace(/\s+/g, " ").trim();
    if (trimmed.length > 60) return `${trimmed.slice(0, 57)}...`;
    return trimmed;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return String(value);
    // Keep stable precision without noise.
    return Number.isInteger(value) ? String(value) : value.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
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
  let bestRow = null;
  let bestStringish = 0;
  let bestNonEmpty = 0;
  for (let r = 0; r < maxRowsToCheck; r += 1) {
    const row = cells[r] || [];
    const nonEmpty = row.filter((c) => c && c.v != null && String(c.v).trim() !== "").length;
    if (nonEmpty === 0) continue;
    const stringish = row.filter((c) => c && typeof c.v === "string" && c.v.trim() !== "").length;
    if (stringish / nonEmpty < 0.6) continue;
    if (stringish > bestStringish || (stringish === bestStringish && nonEmpty > bestNonEmpty)) {
      bestRow = r;
      bestStringish = stringish;
      bestNonEmpty = nonEmpty;
    }
  }
  return bestRow;
}

function countFormulas(cells) {
  let count = 0;
  for (const row of cells) {
    for (const cell of row) {
      if (cell && cell.f != null && String(cell.f).trim() !== "") count += 1;
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

  const lines = [];
  lines.push(
    `${chunk.kind.toUpperCase()}: ${chunk.title} (sheet="${chunk.sheetName}", range="${rectA1}", size=${fullRowCount}x${fullColCount}, formulas≈${formulaCount})`
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
          return h || `Column${c + 1}`;
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
    const maxPreRows = Math.min(headerRow, 2);
    for (let r = 0; r < maxPreRows; r += 1) {
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
        const cell = cells[r][c] || {};
        if (headerNames) {
          const header = headerNames[c] || `Column${c + 1}`;
          if (cell.f) {
            const formula = formatScalar(cell.f);
            const value = formatScalar(cell.v);
            row.push(`${header}(${formula})=${value}`);
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
