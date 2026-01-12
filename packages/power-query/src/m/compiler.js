import { parseM } from "./parser.js";
import { MLanguageCompileError } from "./errors.js";
import { valueKey } from "../valueKey.js";
import { MS_PER_DAY, PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../values.js";
import {
  SOURCE_FUNCTIONS,
  TABLE_FUNCTIONS,
  constantIdentifierValue,
  defaultStepName,
  identifierPartsToName,
  identifierToDataType,
  listAggregationFromIdentifier,
  mTypeNameToDataType,
} from "./stdlib.js";

/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QuerySource} QuerySource
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").FilterPredicate} FilterPredicate
 * @typedef {import("../model.js").ComparisonPredicate} ComparisonPredicate
 * @typedef {import("../model.js").SortSpec} SortSpec
 * @typedef {import("../model.js").Aggregation} Aggregation
 * @typedef {import("../model.js").DataType} DataType
 *
 * @typedef {import("./ast.js").MProgram} MProgram
 * @typedef {import("./ast.js").MExpression} MExpression
 * @typedef {import("./ast.js").MLetExpression} MLetExpression
 * @typedef {import("./ast.js").MIdentifier} MIdentifier
 * @typedef {import("./ast.js").MCallExpression} MCallExpression
 * @typedef {import("./ast.js").MFieldAccessExpression} MFieldAccessExpression
 * @typedef {import("./ast.js").MItemAccessExpression} MItemAccessExpression
 * @typedef {import("./ast.js").MRecordExpression} MRecordExpression
 * @typedef {import("./ast.js").MListExpression} MListExpression
 * @typedef {import("./errors.js").MLocation} MLocation
 */

/**
 * @typedef {{
 *   id?: string;
 *   name?: string;
 *   // Optional schema information used for column validation.
 *   // - `tables` can be used to validate `Excel.CurrentWorkbook` references.
 *   // - `initialSchema` can seed validation for scripts without explicit sources.
 *   tables?: Record<string, { columns: string[] }>;
 *   initialSchema?: string[] | null;
 * }} CompileOptions
 */

/**
 * @typedef {{
 *   source: QuerySource;
 *   steps: QueryStep[];
 *   schema: string[] | null;
 * }} Pipeline
 */

/**
 * @typedef {{ kind: "pipeline"; pipeline: Pipeline } | { kind: "value"; value: unknown }} BindingValue
 */

class CompilerContext {
  /**
   * @param {string} sourceText
   * @param {CompileOptions} options
   */
  constructor(sourceText, options) {
    this.sourceText = sourceText;
    this.options = options;
    /** @type {Map<string, BindingValue>} */
    this.env = new Map();
    this.stepIndex = 0;
  }

  /**
   * @param {MExpression} node
   * @param {string} message
   * @returns {never}
   */
  error(node, message) {
    throw new MLanguageCompileError(message, {
      location: node.span.start,
      source: this.sourceText,
      found: null,
    });
  }

  /**
   * @param {string} name
   * @param {QueryOperation} operation
   * @returns {QueryStep}
   */
  makeStep(name, operation) {
    this.stepIndex += 1;
    return { id: `s${this.stepIndex}_${operation.type}`, name, operation };
  }
}

/**
 * @param {MExpression} expr
 * @returns {expr is MIdentifier}
 */
function isIdentifier(expr) {
  return expr.type === "Identifier";
}

/**
 * @param {MExpression} expr
 * @returns {expr is MCallExpression}
 */
function isCall(expr) {
  return expr.type === "CallExpression";
}

/**
 * @param {MExpression} expr
 * @returns {expr is MListExpression}
 */
function isList(expr) {
  return expr.type === "ListExpression";
}

/**
 * @param {MExpression} expr
 * @returns {expr is MRecordExpression}
 */
function isRecord(expr) {
  return expr.type === "RecordExpression";
}

/**
 * @param {MExpression} expr
 * @returns {string | null}
 */
function calleeName(expr) {
  if (!isIdentifier(expr)) return null;
  return identifierPartsToName(expr.parts);
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {BindingValue}
 */
function compileExpression(ctx, expr, preferredStepName = null) {
  // Special-case: Excel.CurrentWorkbook(){[Name="X"]}[Content]
  const workbookName = matchExcelCurrentWorkbookSelection(expr);
  if (workbookName) {
    return { kind: "pipeline", pipeline: pipelineFromTableSource(workbookName, ctx.options.tables?.[workbookName]?.columns ?? null) };
  }

  switch (expr.type) {
    case "LetExpression":
      return { kind: "pipeline", pipeline: compileLet(ctx, expr) };
    case "Identifier": {
      const name = identifierPartsToName(expr.parts);
      const value = ctx.env.get(name);
      if (value) return value;
      const constant = constantIdentifierValue(name);
      if (constant !== undefined) return { kind: "value", value: constant };
      ctx.error(expr, `Unknown identifier '${name}'`);
    }
    case "CallExpression":
      return compileCall(ctx, expr, preferredStepName);
    case "Literal":
      return { kind: "value", value: expr.value };
    case "ListExpression":
    case "RecordExpression":
    case "ParenthesizedExpression":
    case "TypeExpression":
    case "FieldAccessExpression":
    case "ItemAccessExpression":
    case "EachExpression":
    case "IfExpression":
    case "FunctionExpression":
    case "TryExpression":
    case "AsExpression":
    case "UnaryExpression":
    case "BinaryExpression": {
      // These can be constants in some contexts; allow evaluation when explicitly requested.
      return { kind: "value", value: evaluateConstant(ctx, expr) };
    }
    default: {
      /** @type {never} */
      const exhausted = expr;
      ctx.error(expr, `Unsupported expression type '${exhausted.type}'`);
    }
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MLetExpression} expr
 * @returns {Pipeline}
 */
function compileLet(ctx, expr) {
  for (const binding of expr.bindings) {
    const compiled = compileExpression(ctx, binding.value, binding.name.name);
    ctx.env.set(binding.name.name, compiled);
  }
  const out = compileExpression(ctx, expr.body);
  if (out.kind !== "pipeline") {
    ctx.error(expr.body, "The 'in' expression must evaluate to a table");
  }
  return out.pipeline;
}

/**
 * @param {CompilerContext} ctx
 * @param {MCallExpression} expr
 * @param {string | null} preferredStepName
 * @returns {BindingValue}
 */
function compileCall(ctx, expr, preferredStepName) {
  const name = calleeName(expr.callee);
  if (!name) ctx.error(expr, "Unsupported call target");

  if (name === "Table.ColumnNames") {
    const tableArg = expr.args[0];
    if (!tableArg) ctx.error(expr, "Table.ColumnNames requires a table argument");
    const input = compileExpression(ctx, tableArg);
    if (input.kind !== "pipeline") ctx.error(tableArg, "Expected a table value");
    const schema = input.pipeline.schema;
    if (!schema) {
      ctx.error(expr, "Table.ColumnNames requires a known schema in this subset");
    }
    return { kind: "value", value: schema.slice() };
  }

  if (name === "Table.RowCount") {
    const tableArg = expr.args[0];
    if (!tableArg) ctx.error(expr, "Table.RowCount requires a table argument");
    const input = compileExpression(ctx, tableArg);
    if (input.kind !== "pipeline") ctx.error(tableArg, "Expected a table value");
    const pipeline = input.pipeline;
    if (pipeline.steps.length > 0 || pipeline.source.type !== "range") {
      ctx.error(expr, "Table.RowCount is only supported for Range.FromValues sources in this subset");
    }
    const values = pipeline.source.range.values;
    const hasHeaders = pipeline.source.range.hasHeaders ?? true;
    const count = Math.max(0, values.length - (hasHeaders ? 1 : 0));
    return { kind: "value", value: count };
  }

  if (name === "Table.Combine") {
    return { kind: "pipeline", pipeline: compileTableCombineCall(ctx, expr, preferredStepName) };
  }

  if (TABLE_FUNCTIONS.has(name)) {
    return { kind: "pipeline", pipeline: compileTableFunctionCall(ctx, name, expr, preferredStepName) };
  }

  if (SOURCE_FUNCTIONS.has(name)) {
    return { kind: "pipeline", pipeline: compileSourceFunctionCall(ctx, name, expr) };
  }

  // Value-only functions used as literals.
  return { kind: "value", value: evaluateCallConstant(ctx, name, expr) };
}

/**
 * @param {CompilerContext} ctx
 * @param {string} name
 * @param {MCallExpression} expr
 * @returns {Pipeline}
 */
function compileSourceFunctionCall(ctx, name, expr) {
  switch (name) {
    case "Range.FromValues": {
      const rowsArg = expr.args[0];
      if (!rowsArg) ctx.error(expr, "Range.FromValues requires a list argument");
      const grid = evaluateConstant(ctx, rowsArg);
      if (!Array.isArray(grid) || !grid.every((r) => Array.isArray(r))) {
        ctx.error(rowsArg, "Range.FromValues expects a list of rows (list of lists)");
      }
      const options = expr.args[1] ? evaluateRecordOptions(ctx, expr.args[1]) : {};
      const hasHeaders = options.hasHeaders ?? true;
      /** @type {QuerySource} */
      const source = { type: "range", range: { values: /** @type {unknown[][]} */ (grid), hasHeaders } };
      const schema = hasHeaders ? inferSchemaFromGrid(/** @type {unknown[][]} */ (grid)) : null;
      return { source, steps: [], schema };
    }
    case "Excel.CurrentWorkbook": {
      // We support a convenient subset:
      //  - Excel.CurrentWorkbook("TableName")
      //  - Excel.CurrentWorkbook(){[Name="TableName"]}[Content] (handled earlier)
      const tableNameExpr = expr.args[0];
      if (!tableNameExpr) ctx.error(expr, "Excel.CurrentWorkbook requires a table name in this subset");
      const tableName = expectText(ctx, tableNameExpr);
      return pipelineFromTableSource(tableName, ctx.options.tables?.[tableName]?.columns ?? null);
    }
    case "Csv.Document": {
      const path = compileFilePathArg(ctx, expr.args[0], "Csv.Document");
      const options = expr.args[1] ? evaluateRecordOptions(ctx, expr.args[1]) : {};
      /** @type {QuerySource} */
      const source = {
        type: "csv",
        path,
        options: {
          delimiter: typeof options.delimiter === "string" ? options.delimiter : undefined,
          hasHeaders: typeof options.hasHeaders === "boolean" ? options.hasHeaders : undefined,
        },
      };
      return { source, steps: [], schema: null };
    }
    case "Json.Document": {
      const path = compileFilePathArg(ctx, expr.args[0], "Json.Document");
      const jsonPathRaw = expr.args[1] ? evaluateConstant(ctx, expr.args[1]) : undefined;
      const jsonPath = typeof jsonPathRaw === "string" ? jsonPathRaw : undefined;
      /** @type {QuerySource} */
      const source = { type: "json", path, jsonPath };
      return { source, steps: [], schema: null };
    }
    case "Web.Contents": {
      const urlExpr = expr.args[0];
      if (!urlExpr) ctx.error(expr, "Web.Contents requires a URL");
      const url = expectText(ctx, urlExpr);
      const options = expr.args[1] ? evaluateRecordOptions(ctx, expr.args[1]) : {};
      const headersRaw = options.headers;
      const headers = headersRaw && typeof headersRaw === "object" && !Array.isArray(headersRaw) ? headersRaw : undefined;
      const method = typeof options.method === "string" ? options.method : "GET";
      /** @type {QuerySource} */
      const source = { type: "api", url, method, headers: /** @type {any} */ (headers) };
      return { source, steps: [], schema: null };
    }
    case "OData.Feed": {
      const urlExpr = expr.args[0];
      if (!urlExpr) ctx.error(expr, "OData.Feed requires a URL");
      const url = expectText(ctx, urlExpr);
      const options = expr.args[1] ? evaluateRecordOptions(ctx, expr.args[1]) : {};

      const headersRaw = options.headers;
      const headers = headersRaw && typeof headersRaw === "object" && !Array.isArray(headersRaw) ? headersRaw : undefined;

      /**
       * @param {unknown} input
       * @returns {import("../model.js").APIQuerySource["auth"] | undefined}
       */
      const parseAuth = (input) => {
        if (!input || typeof input !== "object" || Array.isArray(input)) return undefined;
        /** @type {any} */
        const record = input;
        /** @type {any} */
        const normalized = {};
        for (const [k, v] of Object.entries(record)) {
          const key = k.toLowerCase();
          if (key === "type") normalized.type = v;
          if (key === "providerid") normalized.providerId = v;
          if (key === "scopes") normalized.scopes = v;
        }
        if (String(normalized.type ?? "").toLowerCase() !== "oauth2") return undefined;
        if (typeof normalized.providerId !== "string" || !normalized.providerId) return undefined;
        return { type: "oauth2", providerId: normalized.providerId, scopes: normalized.scopes };
      };

      const auth = parseAuth(options.auth);
      const rowsPath = typeof options.rowsPath === "string" ? options.rowsPath : undefined;
      const jsonPath = typeof options.jsonPath === "string" ? options.jsonPath : undefined;

      /** @type {QuerySource} */
      const source = {
        type: "odata",
        url,
        headers: /** @type {any} */ (headers),
        auth,
        ...(rowsPath ? { rowsPath } : {}),
        ...(jsonPath ? { jsonPath } : {}),
      };
      return { source, steps: [], schema: null };
    }
    case "SharePoint.Contents":
    case "SharePoint.Files": {
      const urlExpr = expr.args[0];
      if (!urlExpr) ctx.error(expr, `${name} requires a site URL`);
      const rawSiteUrl = expectText(ctx, urlExpr);
      const siteUrl = normalizeSharePointSiteUrl(ctx, urlExpr, rawSiteUrl);

      const rawOptions = expr.args[1] ? evaluateConstant(ctx, expr.args[1]) : undefined;
      /** @type {any} */
      const parsedOptions = rawOptions && typeof rawOptions === "object" && !Array.isArray(rawOptions) ? rawOptions : null;

      /** @type {any} */
      const outOptions = {};
      /** @type {any} */
      let authValue = undefined;
      /** @type {string | null} */
      let providerId = null;
      /** @type {string[] | string | undefined} */
      let scopes = undefined;

      /**
       * @param {unknown} value
       * @returns {{ type: "oauth2"; providerId: string; scopes?: string[] | string } | null}
       */
      const parseAuth = (value) => {
        if (value == null) return null;
        if (typeof value !== "object" || Array.isArray(value)) return null;
        /** @type {any} */
        const record = value;
        /** @type {string | null} */
        let type = null;
        /** @type {string | null} */
        let pid = null;
        /** @type {string[] | string | undefined} */
        let sc = undefined;
        for (const [k, v] of Object.entries(record)) {
          const key = k.toLowerCase();
          if (key === "type" && typeof v === "string") type = v.toLowerCase();
          if (key === "providerid" && typeof v === "string") pid = v;
          if (key === "scopes") {
            if (typeof v === "string") sc = v;
            if (Array.isArray(v)) sc = v.filter((s) => typeof s === "string");
          }
        }
        if (type && type !== "oauth2") return null;
        if (!pid) return null;
        return sc ? { type: "oauth2", providerId: pid, scopes: sc } : { type: "oauth2", providerId: pid };
      };

      if (parsedOptions) {
        for (const [k, v] of Object.entries(parsedOptions)) {
          const key = k.toLowerCase();
          if (key === "recursive" && typeof v === "boolean") outOptions.recursive = v;
          if (key === "includecontent" && typeof v === "boolean") outOptions.includeContent = v;
          if (key === "auth") authValue = v;
          if (key === "providerid" && typeof v === "string") providerId = v;
          if (key === "scopes") {
            if (typeof v === "string") scopes = v;
            if (Array.isArray(v)) scopes = v.filter((s) => typeof s === "string");
          }
        }
      }

      const authParsed = authValue !== undefined ? parseAuth(authValue) : null;
      if (authValue !== undefined) {
        outOptions.auth = authParsed;
      } else if (providerId) {
        outOptions.auth = scopes ? { type: "oauth2", providerId, scopes } : { type: "oauth2", providerId };
      }

      const options = Object.keys(outOptions).length > 0 ? outOptions : undefined;
      /** @type {QuerySource} */
      const source = { type: "sharepoint", siteUrl, mode: name === "SharePoint.Files" ? "files" : "contents", options };
      return { source, steps: [], schema: null };
    }
    case "Odbc.Query": {
      const connExpr = expr.args[0];
      const queryExpr = expr.args[1];
      if (!connExpr || !queryExpr) ctx.error(expr, "Odbc.Query requires (connectionString, query)");
      const connectionString = expectText(ctx, connExpr);
      const query = expectText(ctx, queryExpr);
      let dialect = undefined;
      try {
        const match = connectionString.match(/\bdriver\s*=\s*\{?([^;}]*)\}?/i);
        const driver = match?.[1] ? String(match[1]).toLowerCase() : "";
        if (driver.includes("postgres")) dialect = "postgres";
        else if (driver.includes("sqlite")) dialect = "sqlite";
      } catch {
        // ignore
      }
      /** @type {QuerySource} */
      const source = { type: "database", connection: { kind: "odbc", connectionString }, query, ...(dialect ? { dialect } : {}) };
      return { source, steps: [], schema: null };
    }
    case "PostgreSQL.Database": {
      const serverExpr = expr.args[0];
      const dbExpr = expr.args[1];
      if (!serverExpr || !dbExpr) ctx.error(expr, "PostgreSQL.Database requires (server, database, query)");
      const serverRaw = expectText(ctx, serverExpr);
      const database = expectText(ctx, dbExpr);
      let query = null;
      const third = expr.args[2];
      if (third) {
        if (isRecord(third)) {
          const opts = evaluateRecordOptions(ctx, third);
          if (typeof opts.query === "string") query = opts.query;
        } else {
          query = expectText(ctx, third);
        }
      }
      if (!query) ctx.error(expr, "PostgreSQL.Database requires a query string in this subset");

      // Power Query typically encodes the server as `host` or `host:port`.
      let host = serverRaw;
      /** @type {number | undefined} */
      let port;
      const match = serverRaw.match(/^(.*):(\d+)$/);
      if (match) {
        host = match[1];
        const parsedPort = Number(match[2]);
        if (Number.isFinite(parsedPort) && parsedPort > 0) port = parsedPort;
      }

      /** @type {QuerySource} */
      const source = {
        type: "database",
        connection: { kind: "postgres", host, ...(port != null ? { port } : {}), database },
        query,
        dialect: "postgres",
      };
      return { source, steps: [], schema: null };
    }
    case "SQLite.Database":
    case "Sqlite.Database": {
      const path = compileFilePathArg(ctx, expr.args[0], name);
      let query = null;
      const second = expr.args[1];
      if (second) {
        if (isRecord(second)) {
          const opts = evaluateRecordOptions(ctx, second);
          if (typeof opts.query === "string") query = opts.query;
        } else {
          query = expectText(ctx, second);
        }
      }
      if (!query) ctx.error(expr, `${name} requires a query string in this subset`);
      /** @type {QuerySource} */
      const source = { type: "database", connection: { kind: "sqlite", path }, query, dialect: "sqlite" };
      return { source, steps: [], schema: null };
    }
    case "Sql.Database": {
      const serverExpr = expr.args[0];
      const dbExpr = expr.args[1];
      if (!serverExpr || !dbExpr) ctx.error(expr, "Sql.Database requires (server, database, query)");
      const server = expectText(ctx, serverExpr);
      const database = expectText(ctx, dbExpr);
      let query = null;
      const third = expr.args[2];
      if (third) {
        if (isRecord(third)) {
          const opts = evaluateRecordOptions(ctx, third);
          if (typeof opts.query === "string") query = opts.query;
        } else {
          query = expectText(ctx, third);
        }
      }
      if (!query) ctx.error(expr, "Sql.Database requires a query string in this subset");
      /** @type {QuerySource} */
      const source = { type: "database", connection: { kind: "sql", server, database }, query, dialect: "sqlserver" };
      return { source, steps: [], schema: null };
    }
    case "Query.Reference": {
      const idExpr = expr.args[0];
      if (!idExpr) ctx.error(expr, "Query.Reference requires a query id");
      const queryId = expectText(ctx, idExpr);
      /** @type {QuerySource} */
      const source = { type: "query", queryId };
      return { source, steps: [], schema: null };
    }
    default:
      ctx.error(expr, `Unsupported source function '${name}'`);
  }
}

/**
 * @param {string} tableName
 * @param {string[] | null} schema
 * @returns {Pipeline}
 */
function pipelineFromTableSource(tableName, schema) {
  /** @type {QuerySource} */
  const source = { type: "table", table: tableName };
  return { source, steps: [], schema: schema ?? null };
}

/**
 * @param {CompilerContext} ctx
 * @param {string} fnName
 * @param {MCallExpression} expr
 * @param {string | null} preferredStepName
 * @returns {Pipeline}
 */
function compileTableFunctionCall(ctx, fnName, expr, preferredStepName) {
  const tableArg = expr.args[0];
  if (!tableArg) ctx.error(expr, `${fnName} requires a table as its first argument`);

  const input = compileExpression(ctx, tableArg);
  if (input.kind !== "pipeline") ctx.error(tableArg, "Expected a table value");

  const base = input.pipeline;
  const stepBaseName = preferredStepName ?? defaultStepName(fnName);

  /** @type {{ operations: QueryOperation[]; schema: string[] | null }} */
  const compiled = compileTableOperation(ctx, fnName, expr, base.schema);

  const steps = [];
  const count = compiled.operations.length;
  let schema = base.schema;

  compiled.operations.forEach((operation, idx) => {
    const name = count === 1 ? stepBaseName : idx === count - 1 ? stepBaseName : `${stepBaseName} (${idx + 1})`;
    steps.push(ctx.makeStep(name, operation));
    schema = applySchemaAfterOperation(ctx, expr, schema, operation);
  });

  return { source: base.source, steps: [...base.steps, ...steps], schema };
}

/**
 * `Table.Combine({tbl1, tbl2, ...})`
 *
 * This maps to our internal `append` operation, which expects additional
 * tables to come from query references.
 *
 * @param {CompilerContext} ctx
 * @param {MCallExpression} expr
 * @param {string | null} preferredStepName
 * @returns {Pipeline}
 */
function compileTableCombineCall(ctx, expr, preferredStepName) {
  const listExpr = expr.args[0];
  if (!listExpr || listExpr.type !== "ListExpression" || listExpr.elements.length === 0) {
    ctx.error(expr, "Table.Combine expects a non-empty list of tables");
  }

  const first = compileExpression(ctx, listExpr.elements[0]);
  if (first.kind !== "pipeline") ctx.error(listExpr.elements[0], "Table.Combine list elements must be tables");

  const queryIds = [];
  for (const element of listExpr.elements.slice(1)) {
    const compiled = compileExpression(ctx, element);
    if (compiled.kind !== "pipeline") ctx.error(element, "Table.Combine list elements must be tables");
    if (compiled.pipeline.source.type !== "query" || compiled.pipeline.steps.length > 0) {
      ctx.error(element, "In this subset, Table.Combine can only append Query.Reference(...) tables");
    }
    queryIds.push(compiled.pipeline.source.queryId);
  }

  // Single-table `Table.Combine` is a no-op.
  if (queryIds.length === 0) return first.pipeline;

  const stepName = preferredStepName ?? defaultStepName("Table.Combine");
  const operation = { type: "append", queries: queryIds };
  const steps = [...first.pipeline.steps, ctx.makeStep(stepName, operation)];
  return { source: first.pipeline.source, steps, schema: null };
}

/**
 * @param {CompilerContext} ctx
 * @param {string} fnName
 * @param {MCallExpression} expr
 * @param {string[] | null} schema
 * @returns {{ operations: QueryOperation[]; schema: string[] | null }}
 */
function compileTableOperation(ctx, fnName, expr, schema) {
  switch (fnName) {
    case "Table.SelectColumns": {
      const columns = expectTextList(ctx, expr.args[1], "Table.SelectColumns");
      validateColumnsExist(ctx, expr, schema, columns);
      return { operations: [{ type: "selectColumns", columns }], schema: columns };
    }
    case "Table.RemoveColumns": {
      const columns = expectTextList(ctx, expr.args[1], "Table.RemoveColumns");
      validateColumnsExist(ctx, expr, schema, columns);
      const nextSchema = schema ? schema.filter((c) => !columns.includes(c)) : schema;
      return { operations: [{ type: "removeColumns", columns }], schema: nextSchema };
    }
    case "Table.Sort": {
      const sortList = expr.args[1];
      if (!sortList || !isList(sortList)) ctx.error(expr, "Table.Sort expects a list of sort specs");
      /** @type {SortSpec[]} */
      const sortBy = sortList.elements.map((spec) => compileSortSpec(ctx, spec, schema));
      return { operations: [{ type: "sortRows", sortBy }], schema };
    }
    case "Table.SelectRows": {
      const predicateExpr = expr.args[1];
      if (!predicateExpr || predicateExpr.type !== "EachExpression") {
        ctx.error(expr, "Table.SelectRows expects an 'each' predicate");
      }
      const predicate = compilePredicate(ctx, predicateExpr.body, schema);
      return { operations: [{ type: "filterRows", predicate }], schema };
    }
    case "Table.Distinct": {
      const colsExpr = expr.args[1] ?? null;
      const columns = colsExpr ? expectTextList(ctx, colsExpr, "Table.Distinct") : null;
      if (columns) validateColumnsExist(ctx, expr, schema, columns);
      return { operations: [{ type: "distinctRows", columns }], schema };
    }
    case "Table.RemoveRowsWithErrors": {
      const colsExpr = expr.args[1] ?? null;
      const columns = colsExpr ? expectTextList(ctx, colsExpr, "Table.RemoveRowsWithErrors") : null;
      if (columns) validateColumnsExist(ctx, expr, schema, columns);
      return { operations: [{ type: "removeRowsWithErrors", columns }], schema };
    }
    case "Table.Group": {
      const groupCols = expectTextList(ctx, expr.args[1], "Table.Group");
      validateColumnsExist(ctx, expr, schema, groupCols);
      const aggsExpr = expr.args[2];
      if (!aggsExpr || !isList(aggsExpr)) ctx.error(expr, "Table.Group expects a list of aggregations");
      const aggregations = aggsExpr.elements.map((node) => compileAggregation(ctx, node, schema));
      const outSchema = [...groupCols, ...aggregations.map((a) => a.as ?? `${a.op} of ${a.column}`)];
      return { operations: [{ type: "groupBy", groupColumns: groupCols, aggregations }], schema: outSchema };
    }
    case "Table.AddColumn": {
      const nameExpr = expr.args[1];
      const formulaExpr = expr.args[2];
      if (!nameExpr || !formulaExpr) ctx.error(expr, "Table.AddColumn expects (table, columnName, each expr)");
      const name = expectText(ctx, nameExpr);
      let formula = null;
      if (formulaExpr.type === "EachExpression") {
        validateColumnsReferenced(ctx, formulaExpr.body, schema);
        formula = mExpressionToJsFormula(ctx, formulaExpr.body);
      } else if (formulaExpr.type === "Literal" && formulaExpr.literalType === "string") {
        formula = formulaExpr.value;
      } else {
        ctx.error(formulaExpr, "Table.AddColumn expects an 'each' expression or a string formula");
      }
      const nextSchema = schema ? [...schema, name] : schema;
      return { operations: [{ type: "addColumn", name, formula }], schema: nextSchema };
    }
    case "Table.RenameColumns": {
      const pairsExpr = expr.args[1];
      if (!pairsExpr || !isList(pairsExpr)) ctx.error(expr, "Table.RenameColumns expects a list of {old,new} pairs");
      /** @type {QueryOperation[]} */
      const operations = [];
      let nextSchema = schema ? schema.slice() : schema;
      for (const pair of pairsExpr.elements) {
        if (!isList(pair) || pair.elements.length < 2) ctx.error(pair, "Rename pair must be a list: {old, new}");
        const oldName = expectText(ctx, pair.elements[0]);
        const newName = expectText(ctx, pair.elements[1]);
        if (nextSchema) {
          if (!nextSchema.includes(oldName)) ctx.error(pair, `Unknown column '${oldName}'`);
          nextSchema = nextSchema.map((c) => (c === oldName ? newName : c));
        }
        operations.push({ type: "renameColumn", oldName, newName });
      }
      return { operations, schema: nextSchema };
    }
    case "Table.TransformColumnTypes": {
      const typeSpecs = expr.args[1];
      if (!typeSpecs || !isList(typeSpecs)) ctx.error(expr, "Table.TransformColumnTypes expects a list of {column, type} pairs");
      /** @type {QueryOperation[]} */
      const operations = [];
      for (const spec of typeSpecs.elements) {
        if (!isList(spec) || spec.elements.length < 2) ctx.error(spec, "Type spec must be a list: {column, type}");
        const column = expectText(ctx, spec.elements[0]);
        validateColumnsExist(ctx, spec, schema, [column]);
        const dt = compileDataType(ctx, spec.elements[1]);
        operations.push({ type: "changeType", column, newType: dt });
      }
      return { operations, schema };
    }
    case "Table.Pivot": {
      // Common pattern:
      //   Table.Pivot(tbl, List.Distinct(tbl[Attr]), "Attr", "Value", List.Sum)
      // We support:
      //   Table.Pivot(tbl, {"A","B"}, "Attr", "Value", List.Sum)
      //   Table.Pivot(tbl, "Attr", "Value", List.Sum)
      const arg2 = expr.args[1];
      const arg3 = expr.args[2];
      const arg4 = expr.args[3];
      const arg5 = expr.args[4];

      let rowColumnExpr = null;
      let valueColumnExpr = null;
      let aggExpr = null;
      if (arg4) {
        // 4 or 5 args (after table)
        if (arg3 && arg4 && (isList(arg2) || arg2?.type === "CallExpression")) {
          rowColumnExpr = arg3;
          valueColumnExpr = arg4;
          aggExpr = arg5 ?? null;
        } else {
          rowColumnExpr = arg2;
          valueColumnExpr = arg3;
          aggExpr = arg4;
        }
      } else {
        ctx.error(expr, "Table.Pivot expects (table, pivotValues?, pivotColumn, valueColumn, aggregation)");
      }

      const rowColumn = expectText(ctx, rowColumnExpr);
      const valueColumn = expectText(ctx, valueColumnExpr);
      validateColumnsExist(ctx, expr, schema, [rowColumn, valueColumn]);
      const aggregation = compileAggregationOp(ctx, aggExpr);
      return { operations: [{ type: "pivot", rowColumn, valueColumn, aggregation }], schema: null };
    }
    case "Table.Unpivot": {
      const cols = expectTextList(ctx, expr.args[1], "Table.Unpivot");
      validateColumnsExist(ctx, expr, schema, cols);
      const nameColumn = expectText(ctx, expr.args[2]);
      const valueColumn = expectText(ctx, expr.args[3]);
      const nextSchema = schema ? [...schema.filter((c) => !cols.includes(c)), nameColumn, valueColumn] : null;
      return { operations: [{ type: "unpivot", columns: cols, nameColumn, valueColumn }], schema: nextSchema };
    }
    case "Table.FillDown": {
      const cols = expectTextList(ctx, expr.args[1], "Table.FillDown");
      validateColumnsExist(ctx, expr, schema, cols);
      return { operations: [{ type: "fillDown", columns: cols }], schema };
    }
    case "Table.ReplaceValue": {
      const findExpr = expr.args[1];
      const replaceExpr = expr.args[2];
      const columnsExpr = expr.args[4];
      if (!findExpr || !replaceExpr || !columnsExpr) {
        ctx.error(expr, "Table.ReplaceValue expects (table, old, new, replacer, columns)");
      }
      const find = evaluateConstant(ctx, findExpr);
      const replace = evaluateConstant(ctx, replaceExpr);
      const columns = expectTextList(ctx, columnsExpr, "Table.ReplaceValue");
      validateColumnsExist(ctx, expr, schema, columns);
      /** @type {QueryOperation[]} */
      const operations = columns.map((column) => ({ type: "replaceValues", column, find, replace }));
      return { operations, schema };
    }
    case "Table.SplitColumn": {
      const column = expectText(ctx, expr.args[1]);
      validateColumnsExist(ctx, expr, schema, [column]);
      const splitterExpr = expr.args[2];
      if (!splitterExpr) ctx.error(expr, "Table.SplitColumn expects (table, column, delimiter|Splitter...)");
      const delimiter = compileDelimiter(ctx, splitterExpr);
      const newColumnsExpr = expr.args[3] ?? null;
      const newColumns = newColumnsExpr ? expectTextList(ctx, newColumnsExpr, "Table.SplitColumn") : null;
      const operation = newColumns ? { type: "splitColumn", column, delimiter, newColumns } : { type: "splitColumn", column, delimiter };
      return { operations: [operation], schema: newColumns ? schema : null };
    }
    case "Table.TransformColumns": {
      const specsExpr = expr.args[1];
      if (!specsExpr || !isList(specsExpr)) ctx.error(expr, "Table.TransformColumns expects a list of transforms");
      const transforms = specsExpr.elements.map((node) => compileTransformColumnSpec(ctx, node, schema));
      return { operations: [{ type: "transformColumns", transforms }], schema };
    }
    case "Table.PromoteHeaders": {
      return { operations: [{ type: "promoteHeaders" }], schema: null };
    }
    case "Table.DemoteHeaders": {
      return { operations: [{ type: "demoteHeaders" }], schema: null };
    }
    case "Table.ReorderColumns": {
      const columns = expectTextList(ctx, expr.args[1], "Table.ReorderColumns");
      if (new Set(columns).size !== columns.length) {
        ctx.error(expr, "Table.ReorderColumns column order must not contain duplicates");
      }
      const missingExpr = expr.args[2] ?? null;
      const raw = missingExpr ? evaluateConstant(ctx, missingExpr) : "error";
      const missingField = raw === "ignore" || raw === "useNull" || raw === "error" ? raw : "error";
      if (schema && missingField === "error") validateColumnsExist(ctx, expr, schema, columns);
      return { operations: [{ type: "reorderColumns", columns, missingField }], schema };
    }
    case "Table.AddIndexColumn": {
      const nameExpr = expr.args[1];
      if (!nameExpr) ctx.error(expr, "Table.AddIndexColumn expects (table, newColumnName, initialValue?, increment?)");
      const name = expectText(ctx, nameExpr);
      const initialRaw = expr.args[2] ? evaluateConstant(ctx, expr.args[2]) : 0;
      const incrementRaw = expr.args[3] ? evaluateConstant(ctx, expr.args[3]) : 1;
      const initialValue = typeof initialRaw === "number" ? initialRaw : Number(initialRaw);
      const increment = typeof incrementRaw === "number" ? incrementRaw : Number(incrementRaw);
      if (!Number.isFinite(initialValue) || !Number.isFinite(increment)) {
        ctx.error(expr, "Table.AddIndexColumn initialValue and increment must be numbers");
      }
      if (schema && schema.includes(name)) ctx.error(expr, `Column '${name}' already exists`);
      return { operations: [{ type: "addIndexColumn", name, initialValue, increment }], schema };
    }
    case "Table.FirstN": {
      const countExpr = expr.args[1];
      if (!countExpr) ctx.error(expr, "Table.FirstN expects (table, count)");
      const countRaw = evaluateConstant(ctx, countExpr);
      const count = typeof countRaw === "number" ? countRaw : Number(countRaw);
      if (!Number.isFinite(count) || count < 0) ctx.error(expr, "Table.FirstN requires a non-negative count");
      return { operations: [{ type: "take", count }], schema };
    }
    case "Table.Skip": {
      const countExpr = expr.args[1];
      if (!countExpr) ctx.error(expr, "Table.Skip expects (table, count)");
      const countRaw = evaluateConstant(ctx, countExpr);
      const count = typeof countRaw === "number" ? countRaw : Number(countRaw);
      if (!Number.isFinite(count) || count < 0) ctx.error(expr, "Table.Skip requires a non-negative count");
      return { operations: [{ type: "skip", count }], schema };
    }
    case "Table.RemoveRows": {
      const offsetExpr = expr.args[1];
      if (!offsetExpr) ctx.error(expr, "Table.RemoveRows expects (table, offset, count?)");
      const offsetRaw = evaluateConstant(ctx, offsetExpr);
      const offset = typeof offsetRaw === "number" ? offsetRaw : Number(offsetRaw);
      if (!Number.isFinite(offset) || offset < 0) ctx.error(expr, "Table.RemoveRows requires a non-negative offset");
      const countExpr = expr.args[2] ?? null;
      const countRaw = countExpr ? evaluateConstant(ctx, countExpr) : null;
      if (countRaw == null) {
        return { operations: [{ type: "take", count: offset }], schema };
      }
      const count = typeof countRaw === "number" ? countRaw : Number(countRaw);
      if (!Number.isFinite(count) || count < 0) ctx.error(expr, "Table.RemoveRows requires a non-negative count");
      if (offset === 0) {
        return { operations: [{ type: "skip", count }], schema };
      }
      return { operations: [{ type: "removeRows", offset, count }], schema };
    }
    case "Table.CombineColumns": {
      const columns = expectTextList(ctx, expr.args[1], "Table.CombineColumns");
      validateColumnsExist(ctx, expr, schema, columns);
      const combinerExpr = expr.args[2];
      const newNameExpr = expr.args[3];
      if (!combinerExpr || !newNameExpr) {
        ctx.error(expr, "Table.CombineColumns expects (table, columns, combiner, newColumnName)");
      }
      const delimiter = compileCombinerDelimiter(ctx, combinerExpr);
      const newColumnName = expectText(ctx, newNameExpr);
      return { operations: [{ type: "combineColumns", columns, delimiter, newColumnName }], schema };
    }
    case "Table.TransformColumnNames": {
      const fnExpr = expr.args[1];
      if (!fnExpr) ctx.error(expr, "Table.TransformColumnNames expects (table, transformFn)");
      if (fnExpr.type !== "Identifier") {
        ctx.error(fnExpr, "Table.TransformColumnNames only supports built-in Text.* transforms in this subset");
      }
      const fn = identifierPartsToName(fnExpr.parts);
      const transform =
        fn === "Text.Upper" ? "upper" : fn === "Text.Lower" ? "lower" : fn === "Text.Trim" ? "trim" : null;
      if (!transform) {
        ctx.error(fnExpr, `Unsupported Table.TransformColumnNames function '${fn}'`);
      }
      return { operations: [{ type: "transformColumnNames", transform }], schema: null };
    }
    case "Table.ReplaceErrorValues": {
      const specsExpr = expr.args[1];
      if (!specsExpr || !isList(specsExpr)) ctx.error(expr, "Table.ReplaceErrorValues expects a list of {column, value} pairs");
      const replacements = specsExpr.elements.map((spec) => {
        if (!isList(spec) || spec.elements.length < 2) ctx.error(spec, "Replacement spec must be a list: {column, value}");
        const column = expectText(ctx, spec.elements[0]);
        validateColumnsExist(ctx, spec, schema, [column]);
        const value = evaluateConstant(ctx, spec.elements[1]);
        return { column, value };
      });
      return { operations: [{ type: "replaceErrorValues", replacements }], schema };
    }
    case "Table.Join":
    case "Table.NestedJoin":
    case "Table.AddJoinColumn": {
      const isNested = fnName !== "Table.Join";
      const leftKeys = expectJoinKeys(ctx, expr.args[1], fnName);
      validateColumnsExist(ctx, expr, schema, leftKeys);
      const rightTableExpr = expr.args[2];
      if (!rightTableExpr) ctx.error(expr, `${fnName} requires a right table argument`);
      const rightKeys = expectJoinKeys(ctx, expr.args[3], fnName);
      if (leftKeys.length !== rightKeys.length) {
        ctx.error(expr, `${fnName} requires left and right join key lists to have the same length`);
      }
      const newColumnName = isNested ? expectText(ctx, expr.args[4]) : null;
      const joinKindExpr = isNested ? (expr.args[5] ?? null) : (expr.args[4] ?? null);
      const joinType = compileJoinKind(ctx, joinKindExpr);

      // Power Query's full signature supports optional join algorithm + comparer
      // arguments. The join algorithm does not affect results, but we preserve it
      // for round-tripping. When a comparer is provided we apply it to the join
      // key equality semantics.
      const algorithmOrComparerExpr = isNested ? (expr.args[6] ?? null) : (expr.args[5] ?? null);
      const explicitComparerExpr = isNested ? (expr.args[7] ?? null) : (expr.args[6] ?? null);

      let joinAlgorithm = null;
      if (algorithmOrComparerExpr && !isComparerExpr(ctx, algorithmOrComparerExpr)) {
        joinAlgorithm = compileJoinAlgorithm(ctx, algorithmOrComparerExpr);
      }

      let comparerExpr = explicitComparerExpr;
      if (!comparerExpr && algorithmOrComparerExpr && isComparerExpr(ctx, algorithmOrComparerExpr)) {
        comparerExpr = algorithmOrComparerExpr;
      }
      const comparerSpec = comparerExpr ? compileJoinComparerArg(ctx, comparerExpr, leftKeys.length) : null;

      const rightQuery = expectQueryReferenceId(ctx, rightTableExpr, fnName);
      return {
        operations: [
          {
            type: "merge",
            rightQuery,
            joinType,
            leftKeys,
            rightKeys,
            joinMode: isNested ? "nested" : "flat",
            ...(newColumnName != null ? { newColumnName } : null),
            ...(joinAlgorithm != null ? { joinAlgorithm } : null),
            ...(comparerSpec?.comparer != null ? { comparer: comparerSpec.comparer } : null),
            ...(comparerSpec?.comparers != null ? { comparers: comparerSpec.comparers } : null),
          },
        ],
        schema: null,
      };
    }
    case "Table.ExpandTableColumn": {
      const column = expectText(ctx, expr.args[1]);
      validateColumnsExist(ctx, expr, schema, [column]);

      const columnsExpr = expr.args[2] ?? null;
      /** @type {string[] | null} */
      let columns = null;
      if (columnsExpr) {
        const raw = evaluateConstant(ctx, columnsExpr);
        if (raw != null) {
          columns = expectTextList(ctx, columnsExpr, "Table.ExpandTableColumn");
        }
      }

      const newNamesExpr = expr.args[3] ?? null;
      /** @type {string[] | null} */
      let newColumnNames = null;
      if (newNamesExpr) {
        const raw = evaluateConstant(ctx, newNamesExpr);
        if (raw != null) {
          newColumnNames = expectTextList(ctx, newNamesExpr, "Table.ExpandTableColumn");
        }
      }

      if (newColumnNames && columns && newColumnNames.length !== columns.length) {
        ctx.error(expr, "Table.ExpandTableColumn new column names must match the expanded column list length");
      }

      return {
        operations: [{ type: "expandTableColumn", column, columns, newColumnNames }],
        schema: null,
      };
    }
    default:
      ctx.error(expr, `Unsupported table function '${fnName}'`);
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} node
 * @param {string[] | null} schema
 * @param {QueryOperation} operation
 * @returns {string[] | null}
 */
function applySchemaAfterOperation(ctx, node, schema, operation) {
  switch (operation.type) {
    case "selectColumns":
      return operation.columns.slice();
    case "removeColumns":
      return schema ? schema.filter((c) => !operation.columns.includes(c)) : schema;
    case "renameColumn":
      return schema ? schema.map((c) => (c === operation.oldName ? operation.newName : c)) : schema;
    case "addColumn":
      return schema ? [...schema, operation.name] : schema;
    case "addIndexColumn":
      return schema ? [...schema, operation.name] : schema;
    case "reorderColumns": {
      if (!schema) return schema;
      const missingField = operation.missingField ?? "error";
      const seen = new Set();
      /** @type {string[]} */
      const out = [];
      for (const name of operation.columns) {
        if (seen.has(name)) continue;
        seen.add(name);
        if (schema.includes(name)) out.push(name);
        else if (missingField === "useNull") out.push(name);
        else if (missingField === "ignore") continue;
        else return null;
      }
      for (const name of schema) {
        if (seen.has(name)) continue;
        out.push(name);
      }
      return out;
    }
    case "combineColumns": {
      if (!schema) return schema;
      const remove = new Set(operation.columns);
      /** @type {string[]} */
      const out = [];
      let inserted = false;
      for (const name of schema) {
        if (!remove.has(name)) {
          out.push(name);
          continue;
        }
        if (!inserted) {
          out.push(operation.newColumnName);
          inserted = true;
        }
      }
      if (!inserted) out.push(operation.newColumnName);
      return out;
    }
    case "changeType":
    case "filterRows":
    case "sortRows":
    case "distinctRows":
    case "removeRowsWithErrors":
    case "transformColumns":
    case "fillDown":
    case "replaceValues":
    case "replaceErrorValues":
    case "take":
    case "skip":
    case "removeRows":
      return schema;
    case "groupBy":
      return [...operation.groupColumns, ...operation.aggregations.map((a) => a.as ?? `${a.op} of ${a.column}`)];
    case "pivot":
    case "promoteHeaders":
    case "demoteHeaders":
    case "transformColumnNames":
      return null;
    case "splitColumn": {
      if (!schema) return schema;
      const names = operation.newColumns;
      if (!names || names.length === 0) return null;
      const idx = schema.indexOf(operation.column);
      if (idx === -1) return null;
      const next = schema.slice();
      next.splice(idx, 1, ...names);
      return next;
    }
    case "unpivot":
      return schema ? [...schema.filter((c) => !operation.columns.includes(c)), operation.nameColumn, operation.valueColumn] : null;
    default:
      // If the engine can't predict the resulting columns, stop validating downstream columns.
      return null;
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 * @param {string[]} columns
 */
function validateColumnsExist(ctx, expr, schema, columns) {
  if (!schema) return;
  const missing = columns.filter((c) => !schema.includes(c));
  if (missing.length) {
    ctx.error(expr, `Unknown column${missing.length === 1 ? "" : "s"}: ${missing.map((c) => `'${c}'`).join(", ")}`);
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 */
function validateColumnsReferenced(ctx, expr, schema) {
  if (!schema) return;
  for (const col of collectFieldReferences(expr)) {
    if (!schema.includes(col)) ctx.error(expr, `Unknown column '${col}'`);
  }
}

/**
 * @param {MExpression} expr
 * @param {Set<string>} [out]
 * @returns {Set<string>}
 */
function collectFieldReferences(expr, out = new Set()) {
  switch (expr.type) {
    case "FieldAccessExpression":
      if (expr.base == null) out.add(expr.field);
      if (expr.base) collectFieldReferences(expr.base, out);
      return out;
    case "IfExpression":
      collectFieldReferences(expr.test, out);
      collectFieldReferences(expr.consequent, out);
      collectFieldReferences(expr.alternate, out);
      return out;
    case "TryExpression":
      collectFieldReferences(expr.expression, out);
      if (expr.otherwise) collectFieldReferences(expr.otherwise, out);
      return out;
    case "AsExpression":
      collectFieldReferences(expr.expression, out);
      return out;
    case "FunctionExpression":
      collectFieldReferences(expr.body, out);
      return out;
    case "BinaryExpression":
      collectFieldReferences(expr.left, out);
      collectFieldReferences(expr.right, out);
      return out;
    case "UnaryExpression":
      collectFieldReferences(expr.argument, out);
      return out;
    case "CallExpression":
      collectFieldReferences(expr.callee, out);
      expr.args.forEach((a) => collectFieldReferences(a, out));
      return out;
    case "ListExpression":
      expr.elements.forEach((e) => collectFieldReferences(e, out));
      return out;
    case "RecordExpression":
      expr.fields.forEach((f) => collectFieldReferences(f.value, out));
      return out;
    case "ParenthesizedExpression":
      return collectFieldReferences(expr.expression, out);
    case "LetExpression":
      expr.bindings.forEach((b) => collectFieldReferences(b.value, out));
      return collectFieldReferences(expr.body, out);
    default:
      return out;
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | undefined | null} expr
 * @param {string} context
 * @returns {string[]}
 */
function expectTextList(ctx, expr, context) {
  if (!expr) ctx.error({ span: { start: { offset: 0, line: 1, column: 1 }, end: { offset: 0, line: 1, column: 1 } } }, `${context} requires a list`);
  const value = evaluateConstant(ctx, expr);
  if (!Array.isArray(value)) ctx.error(expr, `${context} expects a list`);
  return value.map((v) => (typeof v === "string" ? v : String(v)));
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | undefined | null} expr
 * @returns {string}
 */
function expectText(ctx, expr) {
  if (!expr) ctx.error({ span: { start: { offset: 0, line: 1, column: 1 }, end: { offset: 0, line: 1, column: 1 } } }, "Expected text");
  const value = evaluateConstant(ctx, expr);
  if (typeof value !== "string") ctx.error(expr, "Expected a text value");
  return value;
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 * @returns {SortSpec}
 */
function compileSortSpec(ctx, expr, schema) {
  if (!isList(expr) || expr.elements.length === 0) ctx.error(expr, "Sort spec must be a list like {\"Column\", Order.Descending}");
  const column = expectText(ctx, expr.elements[0]);
  validateColumnsExist(ctx, expr, schema, [column]);
  const dirVal = expr.elements[1] ? evaluateConstant(ctx, expr.elements[1]) : "ascending";
  const direction = dirVal === "descending" ? "descending" : "ascending";
  return { column, direction };
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 * @returns {Aggregation}
 */
function compileAggregation(ctx, expr, schema) {
  if (!isList(expr) || expr.elements.length < 2) {
    ctx.error(expr, "Aggregation must be a list like {\"Total\", each List.Sum([Sales])}");
  }
  const as = expectText(ctx, expr.elements[0]);
  const fnExpr = expr.elements[1];
  if (fnExpr.type !== "EachExpression") ctx.error(fnExpr, "Aggregation must use an 'each' expression");
  const body = fnExpr.body;
  if (!isCall(body)) ctx.error(body, "Aggregation body must be a function call");
  const aggFnName = calleeName(body.callee);
  if (!aggFnName) ctx.error(body, "Unsupported aggregation function");
  const op = listAggregationFromIdentifier(aggFnName);
  if (!op) ctx.error(body, `Unsupported aggregation function '${aggFnName}'`);
  const arg0 = body.args[0];
  if (!arg0 || arg0.type !== "FieldAccessExpression" || arg0.base != null) {
    ctx.error(body, "Aggregation function must be called with a column reference like [Sales]");
  }
  const column = arg0.field;
  validateColumnsExist(ctx, body, schema, [column]);
  return { column, op, as };
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | null} expr
 * @returns {Aggregation["op"]}
 */
function compileAggregationOp(ctx, expr) {
  if (!expr) return "sum";
  if (isIdentifier(expr)) {
    const name = identifierPartsToName(expr.parts);
    const op = listAggregationFromIdentifier(name);
    if (op) return op;
    const constVal = constantIdentifierValue(name);
    if (typeof constVal === "string" && (constVal === "ascending" || constVal === "descending")) {
      // Wrong kind of constant.
      ctx.error(expr, "Expected an aggregation function like List.Sum");
    }
    ctx.error(expr, `Unsupported aggregation function '${name}'`);
  }
  const value = evaluateConstant(ctx, expr);
  if (typeof value === "string") {
    const lower = value.toLowerCase();
    if (lower === "sum") return "sum";
    if (lower === "count") return "count";
    if (lower === "average") return "average";
    if (lower === "min") return "min";
    if (lower === "max") return "max";
    if (lower === "countdistinct") return "countDistinct";
  }
  ctx.error(expr, "Unsupported aggregation");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {DataType}
 */
function compileDataType(ctx, expr) {
  if (expr.type === "TypeExpression") return mTypeNameToDataType(expr.name);
  if (expr.type === "Identifier") {
    const name = identifierPartsToName(expr.parts);
    const dt = identifierToDataType(name);
    if (dt) return dt;
  }
  const value = evaluateConstant(ctx, expr);
  if (typeof value === "string") return mTypeNameToDataType(value);
  ctx.error(expr, "Unsupported type expression");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {string}
 */
function compileDelimiter(ctx, expr) {
  if (expr.type === "Literal" && expr.literalType === "string") return expr.value;
  if (expr.type === "CallExpression") {
    const fn = calleeName(expr.callee);
    if (fn === "Splitter.SplitTextByDelimiter") {
      const delim = expr.args[0];
      return expectText(ctx, delim);
    }
  }
  ctx.error(expr, "Unsupported splitter; expected a delimiter string or Splitter.SplitTextByDelimiter(\";\")");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {string}
 */
function compileCombinerDelimiter(ctx, expr) {
  if (expr.type === "Literal" && expr.literalType === "string") return expr.value;
  if (expr.type === "CallExpression") {
    const fn = calleeName(expr.callee);
    if (fn === "Combiner.CombineTextByDelimiter") {
      const delim = expr.args[0];
      return expectText(ctx, delim);
    }
  }
  ctx.error(expr, 'Unsupported combiner; expected a delimiter string or Combiner.CombineTextByDelimiter(";", QuoteStyle.None)');
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | undefined | null} expr
 * @param {string} fnName
 * @returns {string[]}
 */
function expectJoinKeys(ctx, expr, fnName) {
  if (!expr) ctx.error(exprSpanStart(), `${fnName} requires join key columns`);
  const value = evaluateConstant(ctx, expr);
  if (typeof value === "string") return [value];
  if (Array.isArray(value) && value.length > 0) {
    return value.map((v) => (typeof v === "string" ? v : String(v)));
  }
  ctx.error(expr, `${fnName} join keys must be a column name or a non-empty list of column names`);
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | null} expr
 * @returns {"inner" | "left" | "right" | "full"}
 */
function compileJoinKind(ctx, expr) {
  if (!expr) return "inner";
  if (isIdentifier(expr)) {
    const name = identifierPartsToName(expr.parts);
    switch (name) {
      case "JoinKind.Inner":
        return "inner";
      case "JoinKind.LeftOuter":
        return "left";
      case "JoinKind.RightOuter":
        return "right";
      case "JoinKind.FullOuter":
        return "full";
      default:
        break;
    }
  }
  const value = evaluateConstant(ctx, expr);
  if (value == null) return "inner";
  if (typeof value === "number" && Number.isFinite(value)) {
    switch (Math.trunc(value)) {
      case 0:
        return "inner";
      case 1:
        return "left";
      case 2:
        return "right";
      case 3:
        return "full";
      default:
        break;
    }
  }
  if (typeof value === "string") {
    const lower = value.toLowerCase();
    if (lower === "inner") return "inner";
    if (lower === "left") return "left";
    if (lower === "right") return "right";
    if (lower === "full") return "full";
    if (lower === "leftouter") return "left";
    if (lower === "rightouter") return "right";
    if (lower === "fullouter") return "full";
  }
  ctx.error(expr, "Unsupported join kind (expected JoinKind.Inner/LeftOuter/RightOuter/FullOuter)");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {string | null}
 */
function compileJoinAlgorithm(ctx, expr) {
  const value = evaluateConstant(ctx, expr);
  if (value == null) return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    switch (Math.trunc(value)) {
      case 0:
        return "dynamic";
      case 1:
        return "sortMerge";
      case 2:
        return "leftHash";
      case 3:
        return "rightHash";
      case 4:
        return "pairwiseHash";
      default:
        break;
    }
  }
  if (typeof value === "string") {
    const lower = value.toLowerCase();
    if (lower === "dynamic") return "dynamic";
    if (lower === "sortmerge") return "sortMerge";
    if (lower === "lefthash") return "leftHash";
    if (lower === "righthash") return "rightHash";
    if (lower === "pairwisehash") return "pairwiseHash";
  }
  ctx.error(expr, "Unsupported join algorithm (expected JoinAlgorithm.Dynamic/SortMerge/LeftHash/RightHash/PairwiseHash)");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {boolean}
 */
function isComparerExpr(ctx, expr) {
  if (isIdentifier(expr)) {
    const name = identifierPartsToName(expr.parts);
    if (name.startsWith("Comparer.")) return true;
    if (name.startsWith("JoinAlgorithm.")) return false;
  }
  try {
    const value = evaluateConstant(ctx, expr);
    if (Array.isArray(value)) {
      return value.some(
        (v) =>
          v &&
          typeof v === "object" &&
          !Array.isArray(v) &&
          ("caseSensitive" in v || "comparer" in v || "Comparer" in v),
      );
    }
    return value && typeof value === "object" && !Array.isArray(value) && ("caseSensitive" in value || "comparer" in value || "Comparer" in value);
  } catch {
    return false;
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {{ comparer: string; caseSensitive?: boolean } | null}
 */
/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {unknown} input
 * @returns {{ comparer: string; caseSensitive: boolean } | null}
 */
function parseJoinComparer(ctx, expr, input) {
  if (input == null) return null;
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    ctx.error(expr, "Unsupported join comparer (expected Comparer.Ordinal or Comparer.OrdinalIgnoreCase)");
  }

  /** @type {Record<string, unknown>} */
  const record = /** @type {any} */ (input);
  let comparerName = null;
  /** @type {boolean | undefined} */
  let caseSensitive;

  for (const [k, v] of Object.entries(record)) {
    const key = k.toLowerCase();
    if (key === "comparer") comparerName = typeof v === "string" ? v : String(v);
    if (key === "casesensitive") caseSensitive = Boolean(v);
  }

  if (typeof comparerName !== "string") {
    ctx.error(expr, "Unsupported join comparer (expected Comparer.Ordinal or Comparer.OrdinalIgnoreCase)");
  }

  const normalized = comparerName.trim().toLowerCase();
  if (normalized === "ordinal") {
    if (caseSensitive === false) ctx.error(expr, "Comparer.Ordinal must be case sensitive");
    return { comparer: "ordinal", caseSensitive: true };
  }
  if (normalized === "ordinalignorecase") {
    if (caseSensitive === true) ctx.error(expr, "Comparer.OrdinalIgnoreCase must be case insensitive");
    return { comparer: "ordinalIgnoreCase", caseSensitive: false };
  }

  ctx.error(expr, "Unsupported join comparer (expected Comparer.Ordinal or Comparer.OrdinalIgnoreCase)");
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {number} keyCount
 * @returns {{
 *   comparer?: { comparer: string; caseSensitive: boolean } | null;
 *   comparers?: Array<{ comparer: string; caseSensitive: boolean }> | null;
 * } | null}
 */
function compileJoinComparerArg(ctx, expr, keyCount) {
  const raw = evaluateConstant(ctx, expr);
  if (raw == null) return null;

  const defaultComparer = { comparer: "ordinal", caseSensitive: true };

  if (Array.isArray(raw)) {
    if (raw.length === 0) return null;
    if (raw.length === 1) {
      const parsed = parseJoinComparer(ctx, expr, raw[0]);
      return parsed ? { comparer: parsed } : null;
    }
    if (raw.length !== keyCount) {
      ctx.error(expr, `Join comparer list must have length 1 or match join key count (${keyCount}), got ${raw.length}`);
    }

    const comparers = raw.map((entry) => parseJoinComparer(ctx, expr, entry) ?? defaultComparer);
    const first = comparers[0] ?? defaultComparer;
    const allSame = comparers.every((c) => c.comparer === first.comparer && c.caseSensitive === first.caseSensitive);
    return allSame ? { comparer: first } : { comparers };
  }

  const parsed = parseJoinComparer(ctx, expr, raw);
  return parsed ? { comparer: parsed } : null;
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string} context
 * @returns {string}
 */
function expectQueryReferenceId(ctx, expr, context) {
  const compiled = compileExpression(ctx, expr);
  if (compiled.kind !== "pipeline") ctx.error(expr, `${context} expects a table value`);
  if (compiled.pipeline.source.type !== "query" || compiled.pipeline.steps.length > 0) {
    ctx.error(expr, `${context} expects a Query.Reference(...) table in this subset`);
  }
  return compiled.pipeline.source.queryId;
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 * @returns {{ column: string; formula: string; newType: DataType | null }}
 */
function compileTransformColumnSpec(ctx, expr, schema) {
  if (!isList(expr) || expr.elements.length < 2) {
    ctx.error(expr, "Transform spec must be a list like {\"Column\", each _ * 2, type number}");
  }
  const column = expectText(ctx, expr.elements[0]);
  validateColumnsExist(ctx, expr, schema, [column]);

  const fnExpr = expr.elements[1];
  let formula;
  if (fnExpr.type === "EachExpression") {
    formula = mExpressionToJsValueFormula(ctx, fnExpr.body);
  } else if (fnExpr.type === "Literal" && fnExpr.literalType === "string") {
    formula = fnExpr.value;
  } else {
    ctx.error(fnExpr, "Transform must be an 'each' expression or a string formula");
  }

  const newTypeExpr = expr.elements[2];
  const newType = newTypeExpr ? compileDataType(ctx, newTypeExpr) : null;
  return { column, formula, newType };
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @param {string[] | null} schema
 * @returns {FilterPredicate}
 */
function compilePredicate(ctx, expr, schema) {
  switch (expr.type) {
    case "ParenthesizedExpression":
      return compilePredicate(ctx, expr.expression, schema);
    case "AsExpression":
      return compilePredicate(ctx, expr.expression, schema);
    case "TryExpression":
      // Best-effort: errors are not modeled in predicate compilation yet.
      return compilePredicate(ctx, expr.expression, schema);
    case "Literal":
      if (expr.literalType === "boolean") {
        // Represent boolean constants with empty boolean operators:
        // - `and []` is true (vacuously)
        // - `or []` is false
        return expr.value ? { type: "and", predicates: [] } : { type: "or", predicates: [] };
      }
      ctx.error(expr, "Only logical literals (true/false) are supported in predicates");
    case "IfExpression": {
      const test = compilePredicate(ctx, expr.test, schema);
      const consequent = compilePredicate(ctx, expr.consequent, schema);
      const alternate = compilePredicate(ctx, expr.alternate, schema);
      return {
        type: "or",
        predicates: [
          { type: "and", predicates: [test, consequent] },
          { type: "and", predicates: [{ type: "not", predicate: test }, alternate] },
        ],
      };
    }
    case "BinaryExpression": {
      if (expr.operator === "and" || expr.operator === "or") {
        const left = compilePredicate(ctx, expr.left, schema);
        const right = compilePredicate(ctx, expr.right, schema);
        const key = expr.operator === "and" ? "and" : "or";
        /** @type {FilterPredicate[]} */
        const predicates = [];
        const push = (p) => {
          if (p.type === key) predicates.push(...p.predicates);
          else predicates.push(p);
        };
        push(left);
        push(right);
        return { type: key, predicates };
      }

      const comparison = compileComparison(ctx, expr, schema);
      return comparison;
    }
    case "UnaryExpression":
      if (expr.operator !== "not") ctx.error(expr, `Unsupported unary operator '${expr.operator}' in predicate`);
      return { type: "not", predicate: compilePredicate(ctx, expr.argument, schema) };
    case "CallExpression":
      return compilePredicateCall(ctx, expr, schema);
    default:
      ctx.error(expr, "Unsupported predicate expression (expected comparisons, and/or, or Text.Contains)");
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {import("./ast.js").MBinaryExpression} expr
 * @param {string[] | null} schema
 * @returns {ComparisonPredicate}
 */
function compileComparison(ctx, expr, schema) {
  const op = expr.operator;
  if (!["=", "<>", "<", "<=", ">", ">="].includes(op)) {
    ctx.error(expr, `Unsupported comparison operator '${op}'`);
  }

  const left = tryColumnRef(expr.left);
  const right = tryColumnRef(expr.right);

  if (left && !right) {
    const value = evaluateConstant(ctx, expr.right);
    return comparisonFromParts(ctx, expr, left, op, value, schema);
  }
  if (right && !left) {
    const value = evaluateConstant(ctx, expr.left);
    const flipped = flipComparisonOperator(op);
    return comparisonFromParts(ctx, expr, right, flipped, value, schema);
  }

  ctx.error(expr, "Comparisons must involve a column reference like [Region] = \"East\"");
}

/**
 * @param {string} op
 * @returns {string}
 */
function flipComparisonOperator(op) {
  switch (op) {
    case "<":
      return ">";
    case "<=":
      return ">=";
    case ">":
      return "<";
    case ">=":
      return "<=";
    default:
      return op;
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} node
 * @param {string} column
 * @param {string} op
 * @param {unknown} value
 * @param {string[] | null} schema
 * @returns {ComparisonPredicate}
 */
function comparisonFromParts(ctx, node, column, op, value, schema) {
  validateColumnsExist(ctx, node, schema, [column]);
  if (value == null) {
    if (op === "=") return { type: "comparison", column, operator: "isNull" };
    if (op === "<>") return { type: "comparison", column, operator: "isNotNull" };
    ctx.error(node, "Cannot compare null with ordering operators");
  }

  /** @type {ComparisonPredicate["operator"]} */
  let operator;
  switch (op) {
    case "=":
      operator = "equals";
      break;
    case "<>":
      operator = "notEquals";
      break;
    case "<":
      operator = "lessThan";
      break;
    case "<=":
      operator = "lessThanOrEqual";
      break;
    case ">":
      operator = "greaterThan";
      break;
    case ">=":
      operator = "greaterThanOrEqual";
      break;
    default:
      ctx.error(node, `Unsupported comparison operator '${op}'`);
  }
  return { type: "comparison", column, operator, value };
}

/**
 * @param {MExpression} expr
 * @returns {string | null}
 */
function tryColumnRef(expr) {
  if (expr.type === "FieldAccessExpression" && expr.base == null) return expr.field;
  return null;
}

/**
 * @param {CompilerContext} ctx
 * @param {import("./ast.js").MCallExpression} expr
 * @param {string[] | null} schema
 * @returns {FilterPredicate}
 */
function compilePredicateCall(ctx, expr, schema) {
  const name = calleeName(expr.callee);
  if (name === "Text.Contains" || name === "Text.StartsWith" || name === "Text.EndsWith") {
    const colExpr = expr.args[0];
    const needleExpr = expr.args[1];
    if (!colExpr || !needleExpr) ctx.error(expr, `${name} requires (text, substring)`);
    const column = tryColumnRef(colExpr);
    if (!column) ctx.error(colExpr, `${name} first argument must be a column reference like [Name]`);
    validateColumnsExist(ctx, colExpr, schema, [column]);
    const value = evaluateConstant(ctx, needleExpr);
    /** @type {ComparisonPredicate["operator"]} */
    const operator =
      name === "Text.Contains" ? "contains" : name === "Text.StartsWith" ? "startsWith" : "endsWith";

    let caseSensitive = false;
    const comparerExpr = expr.args[2];
    if (comparerExpr) {
      const c = evaluateConstant(ctx, comparerExpr);
      if (c && typeof c === "object" && "caseSensitive" in c) {
        caseSensitive = Boolean(/** @type {any} */ (c).caseSensitive);
      }
    }

    return { type: "comparison", column, operator, value, caseSensitive };
  }

  ctx.error(expr, `Unsupported predicate function '${name ?? "unknown"}'`);
}

/**
 * Convert a subset of M expressions to a Power Query formula string compatible
 * with `compileRowFormula` (the sandboxed expression engine).
 *
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {string}
 */
function mExpressionToJsFormula(ctx, expr) {
  switch (expr.type) {
    case "ParenthesizedExpression":
      // Parentheses do not affect the resulting JS expression because this
      // printer always emits explicit grouping around operators. Dropping
      // them keeps formulas stable across pretty-print round trips.
      return mExpressionToJsFormula(ctx, expr.expression);
    case "IfExpression": {
      const test = mExpressionToJsFormula(ctx, expr.test);
      const consequent = mExpressionToJsFormula(ctx, expr.consequent);
      const alternate = mExpressionToJsFormula(ctx, expr.alternate);
      return `((${test}) ? (${consequent}) : (${alternate}))`;
    }
    case "TryExpression":
      // Best-effort: our expression engine does not model M error values today,
      // so `try` is treated as transparent for formula compilation.
      return mExpressionToJsFormula(ctx, expr.expression);
    case "AsExpression":
      // Type assertions don't affect the formula subset we evaluate today.
      return mExpressionToJsFormula(ctx, expr.expression);
    case "Literal":
      if (expr.literalType === "string") return JSON.stringify(expr.value);
      if (expr.literalType === "number") return String(expr.value);
      if (expr.literalType === "boolean") return expr.value ? "true" : "false";
      if (expr.literalType === "null") return "null";
      ctx.error(expr, "Date literals are not supported in formulas");
    case "FieldAccessExpression":
      if (expr.base != null) ctx.error(expr, "Only implicit [Column] references are supported in formulas");
      return `[${expr.field}]`;
    case "Identifier":
      // `each` formulas in this subset only support `[Column]` references.
      ctx.error(expr, "Identifiers are not supported in formulas (use [Column] references)");
    case "UnaryExpression": {
      const arg = mExpressionToJsFormula(ctx, expr.argument);
      if (expr.operator === "not") return `(!(${arg}))`;
      return `(${expr.operator}(${arg}))`;
    }
    case "BinaryExpression": {
      const left = mExpressionToJsFormula(ctx, expr.left);
      const right = mExpressionToJsFormula(ctx, expr.right);
      const op = binaryOperatorToJs(expr.operator);
      return `((${left}) ${op} (${right}))`;
    }
    case "CallExpression":
      return mCallExpressionToJsFormula(ctx, expr, mExpressionToJsFormula);
    default:
      ctx.error(expr, "Unsupported expression in formula");
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MCallExpression} expr
 * @param {(ctx: CompilerContext, expr: MExpression) => string} compileArg
 * @returns {string}
 */
function mCallExpressionToJsFormula(ctx, expr, compileArg) {
  const name = calleeName(expr.callee);
  if (!name) ctx.error(expr, "Unsupported call target");

  const arg = (idx) => {
    const node = expr.args[idx];
    if (!node) ctx.error(expr, `Missing argument ${idx + 1} for ${name}`);
    return compileArg(ctx, node);
  };

  switch (name) {
    case "Text.Upper":
      return `text_upper(${arg(0)})`;
    case "Text.Lower":
      return `text_lower(${arg(0)})`;
    case "Text.Trim":
      return `text_trim(${arg(0)})`;
    case "Text.Length":
      return `text_length(${arg(0)})`;
    case "Text.Contains":
      return `text_contains(${arg(0)}, ${arg(1)})`;
    case "Number.Round": {
      const first = arg(0);
      const second = expr.args[1] ? `, ${arg(1)}` : "";
      return `number_round(${first}${second})`;
    }
    case "Date.FromText":
    case "Date.From":
      return `date_from_text(${arg(0)})`;
    case "Date.AddDays":
      return `date_add_days(${arg(0)}, ${arg(1)})`;
    default:
      ctx.error(expr, `Unsupported function '${name}' in formula`);
  }
}

/**
 * Convert a subset of M expressions into a Power Query formula string that will
 * be evaluated against a single value (bound as `_`).
 *
 * This is used for `Table.TransformColumns`.
 *
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {string}
 */
function mExpressionToJsValueFormula(ctx, expr) {
  switch (expr.type) {
    case "ParenthesizedExpression":
      return mExpressionToJsValueFormula(ctx, expr.expression);
    case "IfExpression": {
      const test = mExpressionToJsValueFormula(ctx, expr.test);
      const consequent = mExpressionToJsValueFormula(ctx, expr.consequent);
      const alternate = mExpressionToJsValueFormula(ctx, expr.alternate);
      return `((${test}) ? (${consequent}) : (${alternate}))`;
    }
    case "TryExpression":
      return mExpressionToJsValueFormula(ctx, expr.expression);
    case "AsExpression":
      return mExpressionToJsValueFormula(ctx, expr.expression);
    case "Identifier": {
      const name = identifierPartsToName(expr.parts);
      if (name === "_") return "_";
      ctx.error(expr, `Unsupported identifier '${name}' in value formula (expected _)`);
    }
    case "Literal":
      if (expr.literalType === "string") return JSON.stringify(expr.value);
      if (expr.literalType === "number") return String(expr.value);
      if (expr.literalType === "boolean") return expr.value ? "true" : "false";
      if (expr.literalType === "null") return "null";
      ctx.error(expr, "Date literals are not supported in formulas");
    case "UnaryExpression": {
      const arg = mExpressionToJsValueFormula(ctx, expr.argument);
      if (expr.operator === "not") return `(!(${arg}))`;
      return `(${expr.operator}(${arg}))`;
    }
    case "BinaryExpression": {
      const left = mExpressionToJsValueFormula(ctx, expr.left);
      const right = mExpressionToJsValueFormula(ctx, expr.right);
      const op = binaryOperatorToJs(expr.operator);
      return `((${left}) ${op} (${right}))`;
    }
    case "CallExpression":
      return mCallExpressionToJsFormula(ctx, expr, mExpressionToJsValueFormula);
    case "FieldAccessExpression":
      ctx.error(expr, "Column references are not supported in Table.TransformColumns formulas (use _)");
    default:
      ctx.error(expr, "Unsupported expression in Table.TransformColumns formula");
  }
}

/**
 * @param {string} op
 * @returns {string}
 */
function binaryOperatorToJs(op) {
  switch (op) {
    case "and":
      return "&&";
    case "or":
      return "||";
    case "=":
      return "==";
    case "<>":
      return "!=";
    case "&":
      return "+";
    default:
      return op;
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {unknown}
 */
function evaluateConstant(ctx, expr) {
  switch (expr.type) {
    case "Literal":
      return expr.value;
    case "ParenthesizedExpression":
      return evaluateConstant(ctx, expr.expression);
    case "UnaryExpression": {
      const value = evaluateConstant(ctx, expr.argument);
      switch (expr.operator) {
        case "not":
          return !Boolean(value);
        case "+":
          return typeof value === "number" ? value : Number(value);
        case "-":
          return typeof value === "number" ? -value : -Number(value);
        default:
          ctx.error(expr, `Unsupported unary operator '${expr.operator}' in constant context`);
      }
    }
    case "BinaryExpression": {
      const left = evaluateConstant(ctx, expr.left);
      const right = evaluateConstant(ctx, expr.right);
      switch (expr.operator) {
        case "and":
          return Boolean(left) && Boolean(right);
        case "or":
          return Boolean(left) || Boolean(right);
        case "=":
          return valueKey(left) === valueKey(right);
        case "<>":
          return valueKey(left) !== valueKey(right);
        case "<":
          return /** @type {any} */ (left) < /** @type {any} */ (right);
        case "<=":
          return /** @type {any} */ (left) <= /** @type {any} */ (right);
        case ">":
          return /** @type {any} */ (left) > /** @type {any} */ (right);
        case ">=":
          return /** @type {any} */ (left) >= /** @type {any} */ (right);
        case "+":
          return /** @type {any} */ (left) + /** @type {any} */ (right);
        case "-":
          return /** @type {any} */ (left) - /** @type {any} */ (right);
        case "*":
          return /** @type {any} */ (left) * /** @type {any} */ (right);
        case "/":
          return /** @type {any} */ (left) / /** @type {any} */ (right);
        case "&":
          return String(left ?? "") + String(right ?? "");
        default:
          ctx.error(expr, `Unsupported binary operator '${expr.operator}' in constant context`);
      }
    }
    case "IfExpression": {
      const test = evaluateConstant(ctx, expr.test);
      return test ? evaluateConstant(ctx, expr.consequent) : evaluateConstant(ctx, expr.alternate);
    }
    case "TryExpression": {
      try {
        return evaluateConstant(ctx, expr.expression);
      } catch (err) {
        if (expr.otherwise) return evaluateConstant(ctx, expr.otherwise);
        return null;
      }
    }
    case "AsExpression":
      return evaluateConstant(ctx, expr.expression);
    case "FunctionExpression": {
      if (ctx.sourceText) {
        const start = expr.span.start.offset;
        const end = expr.span.end.offset;
        if (Number.isFinite(start) && Number.isFinite(end) && start >= 0 && end >= start) {
          return ctx.sourceText.slice(start, end);
        }
      }
      return "<function>";
    }
    case "Identifier": {
      const name = identifierPartsToName(expr.parts);
      const bound = ctx.env.get(name);
      if (bound) {
        if (bound.kind === "value") return bound.value;
        ctx.error(expr, `Identifier '${name}' refers to a table, not a value`);
      }
      const constant = constantIdentifierValue(name);
      if (constant !== undefined) return constant;
      // Allow using identifiers as strings in some contexts (e.g., {"A", "B"} vs {A, B}).
      if (expr.parts.length === 1) return expr.parts[0];
      ctx.error(expr, `Unknown identifier '${name}'`);
    }
    case "ListExpression":
      return expr.elements.map((e) => evaluateConstant(ctx, e));
    case "RecordExpression": {
      /** @type {Record<string, unknown>} */
      const out = {};
      for (const field of expr.fields) {
        out[field.key] = evaluateConstant(ctx, field.value);
      }
      return out;
    }
    case "TypeExpression":
      return expr.name;
    case "CallExpression": {
      const name = calleeName(expr.callee);
      if (!name) ctx.error(expr, "Unsupported constant call");
      return evaluateCallConstant(ctx, name, expr);
    }
    default:
      ctx.error(expr, "Expression is not a constant");
  }
}

/**
 * @param {CompilerContext} ctx
 * @param {string} name
 * @param {MCallExpression} expr
 * @returns {unknown}
 */
function evaluateCallConstant(ctx, name, expr) {
  /**
   * @param {string} encoded
   * @returns {Uint8Array}
   */
  const base64ToBytes = (encoded) => {
    if (typeof Buffer !== "undefined") {
      return new Uint8Array(Buffer.from(encoded, "base64"));
    }
    // eslint-disable-next-line no-undef
    const binary = atob(encoded);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
  };

  switch (name) {
    case "File.Contents": {
      const pathExpr = expr.args[0];
      if (!pathExpr) ctx.error(expr, "File.Contents requires a path");
      return expectText(ctx, pathExpr);
    }
    case "#date": {
      const [y, m, d] = expr.args.map((a) => evaluateConstant(ctx, a));
      if (![y, m, d].every((n) => typeof n === "number")) ctx.error(expr, "#date requires numeric arguments");
      const dt = new Date(Date.UTC(/** @type {number} */ (y), /** @type {number} */ (m) - 1, /** @type {number} */ (d)));
      return dt;
    }
    case "#datetime": {
      const nums = expr.args.map((a) => evaluateConstant(ctx, a));
      if (!nums.every((n) => typeof n === "number")) ctx.error(expr, "#datetime requires numeric arguments");
      const [y, mo, d, hh = 0, mm = 0, ss = 0] = /** @type {number[]} */ (nums);

      // Power Query's `#datetime` allows fractional seconds (e.g. `3.004`).
      // JS Date constructors take a separate millisecond component, so preserve
      // the sub-second precision deterministically.
      const totalMs = Math.round(ss * 1000);
      const wholeSeconds = Math.trunc(totalMs / 1000);
      const millis = totalMs - wholeSeconds * 1000;
      return new Date(Date.UTC(y, mo - 1, d, hh, mm, wholeSeconds, millis));
    }
    case "#time": {
      const nums = expr.args.map((a) => evaluateConstant(ctx, a));
      if (!nums.every((n) => typeof n === "number")) ctx.error(expr, "#time requires numeric arguments");
      const [hh = 0, mm = 0, ss = 0] = /** @type {number[]} */ (nums);
      const totalMs = Math.round(hh * 3_600_000 + mm * 60_000 + ss * 1000);
      return new PqTime(totalMs);
    }
    case "#duration": {
      const nums = expr.args.map((a) => evaluateConstant(ctx, a));
      if (!nums.every((n) => typeof n === "number")) ctx.error(expr, "#duration requires numeric arguments");
      const [days = 0, hh = 0, mm = 0, ss = 0] = /** @type {number[]} */ (nums);
      const totalMs = Math.round(days * MS_PER_DAY + hh * 3_600_000 + mm * 60_000 + ss * 1000);
      return new PqDuration(totalMs);
    }
    case "#datetimezone": {
      const nums = expr.args.map((a) => evaluateConstant(ctx, a));
      if (!nums.every((n) => typeof n === "number")) ctx.error(expr, "#datetimezone requires numeric arguments");
      const [y, mo, d, hh = 0, mm = 0, ss = 0, offH = 0, offM = 0] = /** @type {number[]} */ (nums);

      const offsetMinutes = Math.round(offH * 60 + offM);
      const secMs = Math.round(ss * 1000);
      const wholeSeconds = Math.trunc(secMs / 1000);
      const millis = secMs - wholeSeconds * 1000;
      const utcMs = Date.UTC(y, mo - 1, d, hh, mm, wholeSeconds, millis) - offsetMinutes * 60 * 1000;
      return new PqDateTimeZone(new Date(utcMs), offsetMinutes);
    }
    case "Decimal.FromText": {
      const textExpr = expr.args[0];
      if (!textExpr) ctx.error(expr, "Decimal.FromText requires a text argument");
      const text = evaluateConstant(ctx, textExpr);
      if (typeof text !== "string") ctx.error(expr, "Decimal.FromText requires a text argument");
      return new PqDecimal(text);
    }
    case "Binary.FromText": {
      const textExpr = expr.args[0];
      if (!textExpr) ctx.error(expr, "Binary.FromText requires a text argument");
      const payload = evaluateConstant(ctx, textExpr);
      if (typeof payload !== "string") ctx.error(expr, "Binary.FromText requires a text argument");

      const encodingExpr = expr.args[1] ?? null;
      const encoding = encodingExpr ? evaluateConstant(ctx, encodingExpr) : "base64";
      if (encoding !== "base64") {
        ctx.error(expr, "Binary.FromText only supports BinaryEncoding.Base64 in this subset");
      }
      try {
        return base64ToBytes(payload);
      } catch {
        ctx.error(expr, "Binary.FromText requires valid base64 input");
      }
    }
    default: {
      const constant = constantIdentifierValue(name);
      if (constant !== undefined) return constant;
      ctx.error(expr, `Unsupported function '${name}' in constant context`);
    }
  }
}

/**
 * Normalize common Power Query record option shapes into a simple JS object.
 *
 * @param {CompilerContext} ctx
 * @param {MExpression} expr
 * @returns {{ delimiter?: unknown; hasHeaders?: unknown; headers?: unknown; method?: unknown; query?: unknown; auth?: unknown; jsonPath?: unknown; rowsPath?: unknown }}
 */
function evaluateRecordOptions(ctx, expr) {
  const value = evaluateConstant(ctx, expr);
  if (!value || typeof value !== "object" || Array.isArray(value)) ctx.error(expr, "Expected a record");
  /** @type {any} */
  const record = value;
  /** @type {any} */
  const normalized = {};

  for (const [k, v] of Object.entries(record)) {
    const key = k.toLowerCase();
    if (key === "delimiter") normalized.delimiter = v;
    if (key === "hasheaders") normalized.hasHeaders = v;
    if (key === "method") normalized.method = v;
    if (key === "query") normalized.query = v;
    if (key === "headers") normalized.headers = v;
    if (key === "auth") normalized.auth = v;
    if (key === "jsonpath") normalized.jsonPath = v;
    if (key === "rowspath") normalized.rowsPath = v;
  }

  return normalized;
}

/**
 * @param {unknown[][]} grid
 * @returns {string[] | null}
 */
function inferSchemaFromGrid(grid) {
  const header = grid[0];
  if (!Array.isArray(header)) return null;
  if (!header.every((c) => typeof c === "string")) return null;
  return /** @type {string[]} */ (header.slice());
}

/**
 * Normalize a SharePoint/OneDrive site URL into a canonical HTTPS string.
 *
 * - forces `https:`
 * - strips query/hash
 * - removes a trailing slash (except for `/`)
 *
 * @param {CompilerContext} ctx
 * @param {MExpression} node
 * @param {string} input
 * @returns {string}
 */
function normalizeSharePointSiteUrl(ctx, node, input) {
  let parsed;
  try {
    parsed = new URL(String(input));
  } catch {
    ctx.error(node, "Expected an absolute SharePoint site URL");
  }
  const protocol = parsed.protocol.toLowerCase();
  if (protocol !== "https:" && protocol !== "http:") {
    ctx.error(node, "Expected an http(s) SharePoint site URL");
  }
  const hostname = parsed.hostname.toLowerCase();
  const port = parsed.port && parsed.port !== "443" ? `:${parsed.port}` : "";
  let path = parsed.pathname || "/";
  path = path.replace(/\/{2,}/g, "/");
  if (path.length > 1 && path.endsWith("/")) path = path.slice(0, -1);
  return `https://${hostname}${port}${path}`;
}

/**
 * @param {CompilerContext} ctx
 * @param {MExpression | undefined} arg
 * @param {string} fnName
 * @returns {string}
 */
function compileFilePathArg(ctx, arg, fnName) {
  if (!arg) ctx.error(exprSpanStart(), `${fnName} requires a path or File.Contents(path)`);
  if (arg.type === "Literal" && arg.literalType === "string") return arg.value;
  if (arg.type === "CallExpression" && calleeName(arg.callee) === "File.Contents") {
    const pathExpr = arg.args[0];
    if (!pathExpr) ctx.error(arg, "File.Contents requires a path string");
    return expectText(ctx, pathExpr);
  }
  ctx.error(arg, `${fnName} expects a path string or File.Contents(path)`);
}

/**
 * @returns {import("./ast.js").MExpression}
 */
function exprSpanStart() {
  // Dummy node used when we need a location but don't have a concrete AST node.
  return /** @type {any} */ ({
    span: { start: { offset: 0, line: 1, column: 1 }, end: { offset: 0, line: 1, column: 1 } },
  });
}

/**
 * Detect `Excel.CurrentWorkbook(){[Name="TableName"]}[Content]`.
 *
 * @param {MExpression} expr
 * @returns {string | null}
 */
function matchExcelCurrentWorkbookSelection(expr) {
  if (expr.type !== "FieldAccessExpression") return null;
  if (expr.field !== "Content") return null;
  const base = expr.base;
  if (!base || base.type !== "ItemAccessExpression") return null;
  const item = base;
  if (item.base.type !== "CallExpression") return null;
  const call = item.base;
  const fn = calleeName(call.callee);
  if (fn !== "Excel.CurrentWorkbook") return null;
  if (call.args.length !== 0) return null;
  if (item.key.type !== "RecordExpression") return null;
  const nameField = item.key.fields.find((f) => f.key === "Name" || f.key === "name");
  if (!nameField) return null;
  if (nameField.value.type !== "Literal" || nameField.value.literalType !== "string") return null;
  return nameField.value.value;
}

/**
 * @param {MProgram | string} programOrSource
 * @param {CompileOptions} [options]
 * @returns {Query}
 */
export function compileMToQuery(programOrSource, options = {}) {
  const sourceText = typeof programOrSource === "string" ? programOrSource : "";
  const program = typeof programOrSource === "string" ? parseM(programOrSource) : programOrSource;

  const ctx = new CompilerContext(sourceText, options);
  const compiled = compileExpression(ctx, program.expression);
  if (compiled.kind !== "pipeline") {
    ctx.error(program.expression, "Script must evaluate to a table");
  }

  return {
    id: options.id ?? "m_query",
    name: options.name ?? "M Query",
    source: compiled.pipeline.source,
    steps: compiled.pipeline.steps,
  };
}
