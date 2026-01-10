import { inferValueType, parseScalar } from "../../shared/valueParsing.js";

/**
 * @typedef {"number" | "boolean" | "datetime" | "string"} CsvColumnType
 *
 * @typedef {{
 *   value: any,
 *   format: any
 * }} ParsedCsvCell
 */

/**
 * Infer a column type for each column in the CSV.
 *
 * This is intentionally tolerant of header rows: if the only non-string type present
 * in a column is (say) `"number"`, we classify the column as `"number"` even if the
 * first row is a string header. Individual cells that fail the inferred type fall back
 * to strings.
 *
 * @param {string[][]} rows
 * @param {number} [sampleSize]
 * @returns {CsvColumnType[]}
 */
export function inferColumnTypes(rows, sampleSize = 100) {
  const sample = rows.slice(0, Math.max(0, sampleSize));
  const columnCount = sample.reduce((max, row) => Math.max(max, row.length), 0);

  /** @type {CsvColumnType[]} */
  const types = Array.from({ length: columnCount }, () => "string");

  for (let col = 0; col < columnCount; col++) {
    const values = sample.map((r) => r[col] ?? "").filter((v) => v.trim() !== "");
    if (values.length === 0) {
      types[col] = "string";
      continue;
    }

    const inferred = values.map(inferValueType).filter((t) => t !== "empty");
    const nonStringTypes = new Set(inferred.filter((t) => t !== "string"));

    if (nonStringTypes.size === 0) {
      types[col] = "string";
      continue;
    }

    if (nonStringTypes.size === 1) {
      types[col] = Array.from(nonStringTypes)[0];
      continue;
    }

    types[col] = "string";
  }

  return types;
}

/**
 * @param {string} raw
 * @param {CsvColumnType} type
 * @returns {ParsedCsvCell}
 */
export function parseCellWithColumnType(raw, type) {
  if (raw.trim() === "") return { value: null, format: null };

  if (type === "string") return { value: raw, format: null };

  const parsed = parseScalar(raw);

  if (type === "number" && parsed.type === "number") return { value: parsed.value, format: null };
  if (type === "boolean" && parsed.type === "boolean") return { value: parsed.value, format: null };
  if (type === "datetime" && parsed.type === "datetime") {
    return { value: parsed.value, format: { numberFormat: parsed.numberFormat } };
  }

  // If inference was wrong for this cell, fall back to preserving the string.
  return { value: raw, format: null };
}

