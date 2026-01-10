import { DataTable } from "./table.js";

/**
 * @typedef {import("./model.js").FilterPredicate} FilterPredicate
 */

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function valueToString(value) {
  if (value == null) return "";
  if (isDate(value)) return value.toISOString();
  return String(value);
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @returns {boolean}
 */
function isEqual(a, b) {
  if (a === b) return true;
  if (isDate(a) && isDate(b)) return a.getTime() === b.getTime();
  return false;
}

/**
 * @param {DataTable} table
 * @param {FilterPredicate} predicate
 * @returns {(row: unknown[]) => boolean}
 */
export function compilePredicate(table, predicate) {
  /**
   * @param {unknown[]} row
   * @param {FilterPredicate} node
   * @returns {boolean}
   */
  function evalNode(row, node) {
    switch (node.type) {
      case "and":
        return node.predicates.every((p) => evalNode(row, p));
      case "or":
        return node.predicates.some((p) => evalNode(row, p));
      case "not":
        return !evalNode(row, node.predicate);
      case "comparison": {
        const idx = table.getColumnIndex(node.column);
        const value = row[idx];

        switch (node.operator) {
          case "isNull":
            return value == null;
          case "isNotNull":
            return value != null;
          case "equals":
            return isEqual(value, node.value);
          case "notEquals":
            return !isEqual(value, node.value);
          case "greaterThan":
            return value != null && node.value != null && value > node.value;
          case "greaterThanOrEqual":
            return value != null && node.value != null && value >= node.value;
          case "lessThan":
            return value != null && node.value != null && value < node.value;
          case "lessThanOrEqual":
            return value != null && node.value != null && value <= node.value;
          case "contains": {
            const haystack = valueToString(value);
            const needle = valueToString(node.value);
            if (node.caseSensitive) return haystack.includes(needle);
            return haystack.toLowerCase().includes(needle.toLowerCase());
          }
          case "startsWith": {
            const haystack = valueToString(value);
            const needle = valueToString(node.value);
            if (node.caseSensitive) return haystack.startsWith(needle);
            return haystack.toLowerCase().startsWith(needle.toLowerCase());
          }
          case "endsWith": {
            const haystack = valueToString(value);
            const needle = valueToString(node.value);
            if (node.caseSensitive) return haystack.endsWith(needle);
            return haystack.toLowerCase().endsWith(needle.toLowerCase());
          }
          default: {
            /** @type {never} */
            const exhausted = node.operator;
            throw new Error(`Unsupported operator '${exhausted}'`);
          }
        }
      }
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported predicate type '${exhausted.type}'`);
      }
    }
  }

  return (row) => evalNode(row, predicate);
}

/**
 * SQL generation helpers for query folding.
 */

/**
 * @param {string} identifier
 * @returns {string}
 */
export function quoteIdentifier(identifier) {
  const escaped = identifier.replaceAll('"', '""');
  return `"${escaped}"`;
}

/**
 * @param {unknown} value
 * @returns {string}
 */
export function sqlLiteral(value) {
  if (value == null) return "NULL";
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (isDate(value)) return `'${value.toISOString().replaceAll("'", "''")}'`;
  return `'${String(value).replaceAll("'", "''")}'`;
}

/**
 * @param {FilterPredicate} predicate
 * @param {{ alias?: string }} [options]
 * @returns {string}
 */
export function predicateToSql(predicate, options = {}) {
  const alias = options.alias ?? "t";
  /**
   * @param {FilterPredicate} node
   * @returns {string}
   */
  function toSql(node) {
    switch (node.type) {
      case "and":
        return `(${node.predicates.map(toSql).join(" AND ")})`;
      case "or":
        return `(${node.predicates.map(toSql).join(" OR ")})`;
      case "not":
        return `(NOT ${toSql(node.predicate)})`;
      case "comparison": {
        const colRef = `${alias}.${quoteIdentifier(node.column)}`;
        const caseSensitive = node.caseSensitive ?? false;

        switch (node.operator) {
          case "isNull":
            return `(${colRef} IS NULL)`;
          case "isNotNull":
            return `(${colRef} IS NOT NULL)`;
          case "equals":
            return `(${colRef} = ${sqlLiteral(node.value)})`;
          case "notEquals":
            return `(${colRef} != ${sqlLiteral(node.value)})`;
          case "greaterThan":
            return `(${colRef} > ${sqlLiteral(node.value)})`;
          case "greaterThanOrEqual":
            return `(${colRef} >= ${sqlLiteral(node.value)})`;
          case "lessThan":
            return `(${colRef} < ${sqlLiteral(node.value)})`;
          case "lessThanOrEqual":
            return `(${colRef} <= ${sqlLiteral(node.value)})`;
          case "contains": {
            const pattern = `%${valueToString(node.value).replaceAll("%", "\\%").replaceAll("_", "\\_")}%`;
            if (caseSensitive) return `(${colRef} LIKE ${sqlLiteral(pattern)})`;
            return `(LOWER(${colRef}) LIKE LOWER(${sqlLiteral(pattern)}))`;
          }
          case "startsWith": {
            const pattern = `${valueToString(node.value).replaceAll("%", "\\%").replaceAll("_", "\\_")}%`;
            if (caseSensitive) return `(${colRef} LIKE ${sqlLiteral(pattern)})`;
            return `(LOWER(${colRef}) LIKE LOWER(${sqlLiteral(pattern)}))`;
          }
          case "endsWith": {
            const pattern = `%${valueToString(node.value).replaceAll("%", "\\%").replaceAll("_", "\\_")}`;
            if (caseSensitive) return `(${colRef} LIKE ${sqlLiteral(pattern)})`;
            return `(LOWER(${colRef}) LIKE LOWER(${sqlLiteral(pattern)}))`;
          }
          default: {
            /** @type {never} */
            const exhausted = node.operator;
            throw new Error(`Unsupported operator '${exhausted}'`);
          }
        }
      }
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported predicate type '${exhausted.type}'`);
      }
    }
  }

  return toSql(predicate);
}

