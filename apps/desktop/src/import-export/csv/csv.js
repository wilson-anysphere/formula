/**
 * @typedef {{ delimiter: string }} CsvParseOptions
 * @typedef {{ delimiter: string, newline?: "\n" | "\r\n" }} CsvStringifyOptions
 */

/**
 * RFC4180-ish CSV parser with configurable delimiter.
 *
 * @param {string} text
 * @param {CsvParseOptions} options
 * @returns {string[][]}
 */
export function parseCsv(text, options) {
  const delimiter = options.delimiter;
  if (delimiter.length !== 1) {
    throw new Error(`CSV delimiter must be a single character, got "${delimiter}"`);
  }

  /** @type {string[][]} */
  const rows = [];
  /** @type {string[]} */
  let row = [];
  let field = "";
  let inQuotes = false;

  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];

    if (inQuotes) {
      if (ch === '"') {
        const next = normalized[i + 1];
        if (next === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === delimiter) {
      row.push(field);
      field = "";
      continue;
    }

    if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }

    field += ch;
  }

  row.push(field);
  rows.push(row);

  // Drop the final empty record when the input ends with a newline.
  const last = rows.at(-1);
  if (rows.length > 1 && last && last.length === 1 && last[0] === "" && normalized.endsWith("\n")) {
    rows.pop();
  }

  return rows;
}

function needsQuoting(field, delimiter) {
  return field.includes('"') || field.includes("\n") || field.includes("\r") || field.includes(delimiter);
}

/**
 * @param {string[][]} rows
 * @param {CsvStringifyOptions} options
 * @returns {string}
 */
export function stringifyCsv(rows, options) {
  const delimiter = options.delimiter;
  if (delimiter.length !== 1) {
    throw new Error(`CSV delimiter must be a single character, got "${delimiter}"`);
  }

  const newline = options.newline ?? "\r\n";

  return rows
    .map((row) =>
      row
        .map((field) => {
          if (!needsQuoting(field, delimiter)) return field;
          return `"${field.replaceAll('"', '""')}"`;
        })
        .join(delimiter)
    )
    .join(newline);
}

