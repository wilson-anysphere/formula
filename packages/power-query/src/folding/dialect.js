/**
 * Minimal SQL dialect abstraction used by the query folding engine.
 *
 * The goal is to keep SQL generation conservative + predictable across the
 * connectors we care about (Postgres/MySQL/SQLite) without trying to fully model
 * every vendor-specific syntax variation.
 */

/**
 * @typedef {"postgres" | "mysql" | "sqlite" | "sqlserver"} SqlDialectName
 */

/**
 * @typedef {import("../model.js").SortSpec} SortSpec
 */

/**
 * @typedef {{
 *   name: SqlDialectName;
 *   quoteIdentifier: (identifier: string) => string;
 *   // Cast an arbitrary SQL expression to text for string operations like
 *   // `LIKE`/`LOWER`. This is used to emulate local predicate semantics which
 *   // stringify values before applying `contains`/`startsWith`/`endsWith`.
 *   castText: (sqlExpr: string) => string;
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
 * SQL Server uses `[identifier]` quoting (T-SQL). `]` is escaped as `]]`.
 *
 * @param {string} identifier
 * @returns {string}
 */
function quoteBracket(identifier) {
  const escaped = identifier.replaceAll("]", "]]");
  return `[${escaped}]`;
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
 * SQL Server drivers are typically happy with ISO 8601 date/time strings.
 *
 * Avoid the trailing `Z` timezone suffix because it is not accepted by
 * `DATETIME` / `DATETIME2` string casts (use `DATETIMEOFFSET` if you need a
 * timezone-aware type).
 *
 * @param {Date} date
 * @returns {string}
 */
function formatSqlServerDateTime(date) {
  const iso = date.toISOString();
  return iso.endsWith("Z") ? iso.slice(0, -1) : iso;
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
    case "sqlserver":
      return SQLSERVER_DIALECT;
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
  castText: (expr) => `CAST(${expr} AS TEXT)`,
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
  castText: (expr) => `CAST(${expr} AS CHAR)`,
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
  castText: (expr) => `CAST(${expr} AS TEXT)`,
  formatDateParam: formatIso,
  sortSpecToSql: (alias, spec) => {
    const colRef = `${alias}.${quoteDouble(spec.column)}`;
    const direction = spec.direction === "descending" ? "DESC" : "ASC";
    const nulls = spec.nulls ?? "last";
    const nullFlagDirection = nulls === "first" ? "DESC" : "ASC";
    return [`(${colRef} IS NULL) ${nullFlagDirection}`, `${colRef} ${direction}`];
  },
};

/** @type {SqlDialect} */
export const SQLSERVER_DIALECT = {
  name: "sqlserver",
  quoteIdentifier: quoteBracket,
  castText: (expr) => `CAST(${expr} AS NVARCHAR(MAX))`,
  formatDateParam: formatSqlServerDateTime,
  sortSpecToSql: (alias, spec) => {
    const colRef = `${alias}.${quoteBracket(spec.column)}`;
    const direction = spec.direction === "descending" ? "DESC" : "ASC";
    const nulls = spec.nulls ?? "last";
    const nullFlagDirection = nulls === "first" ? "DESC" : "ASC";
    const nullFlag = `(CASE WHEN ${colRef} IS NULL THEN 1 ELSE 0 END)`;
    return [`${nullFlag} ${nullFlagDirection}`, `${colRef} ${direction}`];
  },
};
