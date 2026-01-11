/**
 * Minimal SQL dialect abstraction used by the query folding engine.
 *
 * The goal is to keep SQL generation conservative + predictable across the
 * connectors we care about (Postgres/MySQL/SQLite) without trying to fully model
 * every vendor-specific syntax variation.
 */

/**
 * @typedef {"postgres" | "mysql" | "sqlite"} SqlDialectName
 */

/**
 * @typedef {import("../model.js").SortSpec} SortSpec
 */

/**
 * @typedef {{
 *   name: SqlDialectName;
 *   quoteIdentifier: (identifier: string) => string;
 *   // Format a JS Date into a string suitable for passing as a SQL parameter.
 *   // (We still parameterize the value; this only normalizes representation.)
 *   formatDateParam: (date: Date) => string;
 *   // Render a sort spec into one or more `ORDER BY` expressions.
 *   // Some dialects need multiple expressions to emulate NULL ordering.
 *   sortSpecToSql: (alias: string, spec: SortSpec) => string[];
 * }} SqlDialect
 */

/**
 * @param {string} identifier
 * @returns {string}
 */
function quoteDouble(identifier) {
  const escaped = identifier.replaceAll('"', '""');
  return `"${escaped}"`;
}

/**
 * @param {string} identifier
 * @returns {string}
 */
function quoteBacktick(identifier) {
  const escaped = identifier.replaceAll("`", "``");
  return `\`${escaped}\``;
}

/**
 * @param {Date} date
 * @returns {string}
 */
function formatIso(date) {
  return date.toISOString();
}

/**
 * MySQL DATETIME literals/params are typically formatted as `YYYY-MM-DD HH:MM:SS`.
 * @param {Date} date
 * @returns {string}
 */
function formatMysqlDateTime(date) {
  const iso = date.toISOString();
  return iso.slice(0, 19).replace("T", " ");
}

/**
 * @param {SqlDialectName} name
 * @returns {SqlDialect}
 */
export function getSqlDialect(name) {
  switch (name) {
    case "postgres":
      return POSTGRES_DIALECT;
    case "mysql":
      return MYSQL_DIALECT;
    case "sqlite":
      return SQLITE_DIALECT;
    default: {
      /** @type {never} */
      const exhausted = name;
      throw new Error(`Unsupported SQL dialect '${exhausted}'`);
    }
  }
}

/** @type {SqlDialect} */
export const POSTGRES_DIALECT = {
  name: "postgres",
  quoteIdentifier: quoteDouble,
  formatDateParam: formatIso,
  sortSpecToSql: (alias, spec) => {
    const colRef = `${alias}.${quoteDouble(spec.column)}`;
    const direction = spec.direction === "descending" ? "DESC" : "ASC";
    const nulls = (spec.nulls ?? "last").toUpperCase();
    return [`${colRef} ${direction} NULLS ${nulls}`];
  },
};

/** @type {SqlDialect} */
export const MYSQL_DIALECT = {
  name: "mysql",
  quoteIdentifier: quoteBacktick,
  formatDateParam: formatMysqlDateTime,
  sortSpecToSql: (alias, spec) => {
    const colRef = `${alias}.${quoteBacktick(spec.column)}`;
    const direction = spec.direction === "descending" ? "DESC" : "ASC";
    const nulls = spec.nulls ?? "last";
    const nullFlagDirection = nulls === "first" ? "DESC" : "ASC";
    return [`(${colRef} IS NULL) ${nullFlagDirection}`, `${colRef} ${direction}`];
  },
};

/** @type {SqlDialect} */
export const SQLITE_DIALECT = {
  name: "sqlite",
  quoteIdentifier: quoteDouble,
  formatDateParam: formatIso,
  sortSpecToSql: (alias, spec) => {
    const colRef = `${alias}.${quoteDouble(spec.column)}`;
    const direction = spec.direction === "descending" ? "DESC" : "ASC";
    const nulls = spec.nulls ?? "last";
    const nullFlagDirection = nulls === "first" ? "DESC" : "ASC";
    return [`(${colRef} IS NULL) ${nullFlagDirection}`, `${colRef} ${direction}`];
  },
};
