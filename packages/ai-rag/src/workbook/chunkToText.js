import { rectToA1 } from "./rect.js";

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
  const row0 = cells[0] || [];
  const nonEmpty = row0.filter((c) => c && c.v != null && String(c.v).trim() !== "").length;
  if (nonEmpty === 0) return null;
  const stringish = row0.filter((c) => c && typeof c.v === "string" && c.v.trim() !== "").length;
  if (stringish / nonEmpty >= 0.6) return 0;
  return null;
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
 * @param {{ sampleRows?: number }} [opts]
 */
export function chunkToText(chunk, opts) {
  const sampleRows = opts?.sampleRows ?? 5;
  const rectA1 = rectToA1(chunk.rect);
  const cells = chunk.cells || [];
  const headerRow = inferHeaderRow(cells);
  const sampledColCount = cells[0]?.length ?? 0;
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

  if (headerRow === 0 && sampledColCount > 0) {
    const headers = [];
    const types = [];
    for (let c = 0; c < sampledColCount; c += 1) {
      const h = formatScalar(cells[0][c]?.v) || `Column${c + 1}`;
      headers.push(h);
      types.push(inferColumnType(cells, c, 0));
    }
    lines.push(`COLUMNS: ${headers.map((h, i) => `${h} (${types[i]})`).join(" | ")}`);
  } else if (sampledColCount > 0) {
    const types = [];
    for (let c = 0; c < sampledColCount; c += 1) types.push(inferColumnType(cells, c, -1));
    lines.push(`COLUMNS: ${types.map((t, i) => `Column${i + 1} (${t})`).join(" | ")}`);
  }

  if (chunk.kind === "formulaRegion") {
    const formulas = [];
    for (let r = 0; r < cells.length && formulas.length < 12; r += 1) {
      for (let c = 0; c < (cells[r]?.length ?? 0) && formulas.length < 12; c += 1) {
        const f = cells[r][c]?.f;
        if (!f) continue;
        formulas.push(String(f).replace(/\s+/g, " ").trim());
      }
    }
    if (formulas.length) lines.push(`FORMULAS: ${formulas.join(" | ")}`);
  } else {
    const startRow = headerRow === 0 ? 1 : 0;
    const rows = [];
    for (let r = startRow; r < Math.min(sampledRowCount, startRow + sampleRows); r += 1) {
      const row = [];
      for (let c = 0; c < sampledColCount; c += 1) {
        const cell = cells[r][c] || {};
        if (cell.f) row.push(formatScalar(cell.f));
        else row.push(formatScalar(cell.v));
      }
      rows.push(row.join(" | "));
    }
    if (rows.length) {
      lines.push("SAMPLE ROWS:");
      for (const row of rows) lines.push(`  - ${row}`);
    }
  }

  return lines.join("\n");
}
