/**
 * Utilities for translating SQL parameter placeholder styles.
 *
 * The SQL folding engine emits `?` placeholders because that keeps query
 * composition simple (nested queries, UNION ALL, joins) and works across many
 * drivers.
 *
 * Some Postgres drivers require `$1..$n` placeholders, so we provide a
 * conservative translator that attempts to only replace `?` that appear in a
 * "value position" (as opposed to Postgres operators like `jsonb ? 'key'`).
 */

const POSTGRES_VALUE_KEYWORDS = new Set([
  "LIKE",
  "IN",
  "NOT",
  "THEN",
  "ELSE",
  "WHEN",
  "LIMIT",
  "OFFSET",
]);

/**
 * @param {string} sql
 * @param {number} paramCount
 * @returns {string}
 */
export function normalizePostgresPlaceholders(sql, paramCount) {
  if (paramCount <= 0) return sql;

  let out = "";
  let replaced = 0;

  let inSingle = false;
  let inDouble = false;
  let inLineComment = false;
  let inBlockComment = false;
  /** @type {string | null} */
  let dollarDelimiter = null;

  for (let i = 0; i < sql.length; i++) {
    const ch = sql[i];
    const next = sql[i + 1] ?? "";

    if (inLineComment) {
      out += ch;
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      out += ch;
      if (ch === "*" && next === "/") {
        out += next;
        i += 1;
        inBlockComment = false;
      }
      continue;
    }

    if (dollarDelimiter) {
      if (sql.startsWith(dollarDelimiter, i)) {
        out += dollarDelimiter;
        i += dollarDelimiter.length - 1;
        dollarDelimiter = null;
      } else {
        out += ch;
      }
      continue;
    }

    if (inSingle) {
      out += ch;
      if (ch === "'") {
        if (next === "'") {
          out += next;
          i += 1;
        } else {
          inSingle = false;
        }
      }
      continue;
    }

    if (inDouble) {
      out += ch;
      if (ch === '"') {
        if (next === '"') {
          out += next;
          i += 1;
        } else {
          inDouble = false;
        }
      }
      continue;
    }

    if (ch === "-" && next === "-") {
      out += ch + next;
      i += 1;
      inLineComment = true;
      continue;
    }

    if (ch === "/" && next === "*") {
      out += ch + next;
      i += 1;
      inBlockComment = true;
      continue;
    }

    if (ch === "'") {
      out += ch;
      inSingle = true;
      continue;
    }

    if (ch === '"') {
      out += ch;
      inDouble = true;
      continue;
    }

    if (ch === "$") {
      const delimiter = parseDollarQuoteDelimiter(sql, i);
      if (delimiter) {
        out += delimiter;
        i += delimiter.length - 1;
        dollarDelimiter = delimiter;
        continue;
      }
    }

    if (ch === "?" && replaced < paramCount && isValuePlaceholder(sql, i)) {
      replaced += 1;
      out += `$${replaced}`;
      continue;
    }

    out += ch;
  }

  if (replaced !== paramCount) {
    throw new Error(`Failed to normalize Postgres SQL placeholders: expected ${paramCount}, replaced ${replaced}`);
  }

  return out;
}

/**
 * @param {string} sql
 * @param {number} pos
 * @returns {string | null}
 */
function parseDollarQuoteDelimiter(sql, pos) {
  const end = sql.indexOf("$", pos + 1);
  if (end === -1) return null;
  const tag = sql.slice(pos + 1, end);
  if (tag !== "" && !/^[A-Za-z0-9_]+$/.test(tag)) return null;
  return sql.slice(pos, end + 1);
}

/**
 * @param {string} sql
 * @param {number} pos
 * @returns {boolean}
 */
function isValuePlaceholder(sql, pos) {
  let i = pos - 1;
  while (i >= 0 && /\s/.test(sql[i])) i -= 1;
  if (i < 0) return true;

  const ch = sql[i];
  if (ch === ")" || ch === "]") return false;
  if (ch === '"' || ch === "'" || /[0-9]/.test(ch)) return false;

  if (/[A-Za-z0-9_]/.test(ch)) {
    let start = i;
    while (start >= 0 && /[A-Za-z0-9_]/.test(sql[start])) start -= 1;
    const word = sql.slice(start + 1, i + 1).toUpperCase();
    return POSTGRES_VALUE_KEYWORDS.has(word);
  }

  return true;
}

