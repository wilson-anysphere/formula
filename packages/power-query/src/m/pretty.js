/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QuerySource} QuerySource
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").FilterPredicate} FilterPredicate
 * @typedef {import("../model.js").ComparisonPredicate} ComparisonPredicate
 * @typedef {import("../model.js").Aggregation} Aggregation
 */

/**
 * @param {string} name
 * @returns {string}
 */
function escapeMString(name) {
  return `"${name.replaceAll('"', '""')}"`;
}

/**
 * @param {string} name
 * @returns {boolean}
 */
function isBareIdentifier(name) {
  return (
    /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) &&
    ![
      "let",
      "in",
      "each",
      "and",
      "or",
      "not",
      "type",
      "if",
      "then",
      "else",
      "try",
      "otherwise",
      "as",
      "true",
      "false",
      "null",
    ].includes(name)
  );
}

/**
 * @param {string} name
 * @returns {string}
 */
function toMIdentifier(name) {
  return isBareIdentifier(name) ? name : `#${escapeMString(name)}`;
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function valueToM(value) {
  if (value == null) return "null";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "null";
  if (typeof value === "string") return escapeMString(value);
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `#date(${value.getUTCFullYear()}, ${value.getUTCMonth() + 1}, ${value.getUTCDate()})`;
  }
  if (Array.isArray(value)) return `{${value.map(valueToM).join(", ")}}`;
  if (typeof value === "object") {
    const entries = Object.entries(value);
    return `[${entries.map(([k, v]) => `${toMIdentifier(k)} = ${valueToM(v)}`).join(", ")}]`;
  }
  return escapeMString(String(value));
}

/**
 * @param {QuerySource} source
 * @returns {string}
 */
function sourceToM(source) {
  switch (source.type) {
    case "table":
      return `Excel.CurrentWorkbook(${escapeMString(source.table)})`;
    case "range": {
      const opts = source.range.hasHeaders != null ? `, [HasHeaders = ${source.range.hasHeaders ? "true" : "false"}]` : "";
      return `Range.FromValues(${valueToM(source.range.values)}${opts})`;
    }
    case "csv": {
      const opts = {};
      if (source.options?.delimiter) opts.Delimiter = source.options.delimiter;
      if (source.options?.hasHeaders != null) opts.HasHeaders = source.options.hasHeaders;
      const optText = Object.keys(opts).length ? `, ${valueToM(opts)}` : "";
      return `Csv.Document(File.Contents(${escapeMString(source.path)})${optText})`;
    }
    case "json": {
      const pathArg = `File.Contents(${escapeMString(source.path)})`;
      const jsonPathArg = source.jsonPath ? `, ${escapeMString(source.jsonPath)}` : "";
      return `Json.Document(${pathArg}${jsonPathArg})`;
    }
    case "api": {
      const opts = {};
      if (source.method && source.method !== "GET") opts.Method = source.method;
      if (source.headers && Object.keys(source.headers).length) opts.Headers = source.headers;
      const optText = Object.keys(opts).length ? `, ${valueToM(opts)}` : "";
      return `Web.Contents(${escapeMString(source.url)}${optText})`;
    }
    case "odata": {
      const opts = {};
      if (source.headers && Object.keys(source.headers).length) opts.Headers = source.headers;
      if (source.auth) opts.Auth = source.auth;
      if (source.rowsPath) opts.RowsPath = source.rowsPath;
      if (source.jsonPath) opts.JsonPath = source.jsonPath;
      const optText = Object.keys(opts).length ? `, ${valueToM(opts)}` : "";
      return `OData.Feed(${escapeMString(source.url)}${optText})`;
    }
    case "database": {
      const c = source.connection;
      if (c && typeof c === "object" && c.kind === "odbc") {
        return `Odbc.Query(${escapeMString(String(c.connectionString ?? ""))}, ${escapeMString(source.query)})`;
      }
      if (c && typeof c === "object" && c.kind === "sql") {
        return `Sql.Database(${escapeMString(String(c.server ?? ""))}, ${escapeMString(String(c.database ?? ""))}, ${escapeMString(source.query)})`;
      }
      return `Odbc.Query(${escapeMString("")}, ${escapeMString(source.query)})`;
    }
    case "query":
      return `Query.Reference(${escapeMString(source.queryId)})`;
    default: {
      /** @type {never} */
      const exhausted = source;
      throw new Error(`Unsupported source type '${exhausted.type}'`);
    }
  }
}

/**
 * @param {FilterPredicate} predicate
 * @returns {string}
 */
function predicateToM(predicate) {
  switch (predicate.type) {
    case "and":
      if (predicate.predicates.length === 0) return "true";
      return predicate.predicates.map(predicateToM).join(" and ");
    case "or":
      if (predicate.predicates.length === 0) return "false";
      return predicate.predicates.map(predicateToM).join(" or ");
    case "not":
      return `not (${predicateToM(predicate.predicate)})`;
    case "comparison":
      return comparisonToM(predicate);
    default: {
      /** @type {never} */
      const exhausted = predicate;
      throw new Error(`Unsupported predicate '${exhausted.type}'`);
    }
  }
}

/**
 * @param {ComparisonPredicate} predicate
 * @returns {string}
 */
function comparisonToM(predicate) {
  const col = `[${predicate.column}]`;
  switch (predicate.operator) {
    case "isNull":
      return `${col} = null`;
    case "isNotNull":
      return `${col} <> null`;
    case "equals":
      return `${col} = ${valueToM(predicate.value)}`;
    case "notEquals":
      return `${col} <> ${valueToM(predicate.value)}`;
    case "greaterThan":
      return `${col} > ${valueToM(predicate.value)}`;
    case "greaterThanOrEqual":
      return `${col} >= ${valueToM(predicate.value)}`;
    case "lessThan":
      return `${col} < ${valueToM(predicate.value)}`;
    case "lessThanOrEqual":
      return `${col} <= ${valueToM(predicate.value)}`;
    case "contains":
      return `Text.Contains(${col}, ${valueToM(predicate.value)})`;
    case "startsWith":
      return `Text.StartsWith(${col}, ${valueToM(predicate.value)})`;
    case "endsWith":
      return `Text.EndsWith(${col}, ${valueToM(predicate.value)})`;
    default: {
      /** @type {never} */
      const exhausted = predicate.operator;
      throw new Error(`Unsupported operator '${exhausted}'`);
    }
  }
}

/**
 * @param {Aggregation["op"]} op
 * @returns {string}
 */
function aggregationOpToM(op) {
  switch (op) {
    case "sum":
      return "List.Sum";
    case "count":
      return "List.Count";
    case "average":
      return "List.Average";
    case "min":
      return "List.Min";
    case "max":
      return "List.Max";
    case "countDistinct":
      return "List.CountDistinct";
    default:
      return "List.Sum";
  }
}

/**
 * @param {import("../model.js").DataType} type
 * @returns {string}
 */
function dataTypeToMTypeExpr(type) {
  switch (type) {
    case "string":
      return "type text";
    case "number":
      return "type number";
    case "boolean":
      return "type logical";
    case "date":
      return "type date";
    case "any":
    default:
      return "type any";
  }
}

/**
 * @param {string} formula
 * @returns {string}
 */
function jsFormulaToM(formula) {
  // Best-effort: convert the restricted JS-ish row/value formula subset back
  // into the M-ish subset this parser understands.
  let expr = formula.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();

  /**
   * @param {string} input
   * @returns {string}
   */
  function stripOuterParens(input) {
    let out = input.trim();
    while (out.startsWith("(") && out.endsWith(")")) {
      let depth = 0;
      let inSingle = false;
      let inDouble = false;
      let wraps = true;
      for (let i = 0; i < out.length; i++) {
        const ch = out[i];
        if (inSingle) {
          if (ch === "'" && out[i - 1] !== "\\") inSingle = false;
          continue;
        }
        if (inDouble) {
          if (ch === '"' && out[i - 1] !== "\\") inDouble = false;
          continue;
        }
        if (ch === "'") {
          inSingle = true;
          continue;
        }
        if (ch === '"') {
          inDouble = true;
          continue;
        }
        if (ch === "(") depth += 1;
        else if (ch === ")") {
          depth -= 1;
          if (depth === 0 && i < out.length - 1) {
            wraps = false;
            break;
          }
        }
      }
      if (!wraps || depth !== 0) break;
      out = out.slice(1, -1).trim();
    }
    return out;
  }

  /**
   * @param {string} input
   * @returns {{ q: number; colon: number } | null}
   */
  function findTopLevelTernary(input) {
    let depth = 0;
    let bracketDepth = 0;
    let inSingle = false;
    let inDouble = false;
    let qIndex = -1;
    let nested = 0;

    for (let i = 0; i < input.length; i++) {
      const ch = input[i];

      if (inSingle) {
        if (ch === "'" && input[i - 1] !== "\\") inSingle = false;
        continue;
      }
      if (inDouble) {
        if (ch === '"' && input[i - 1] !== "\\") inDouble = false;
        continue;
      }
      if (ch === "'") {
        inSingle = true;
        continue;
      }
      if (ch === '"') {
        inDouble = true;
        continue;
      }

      if (ch === "(") depth += 1;
      else if (ch === ")") depth -= 1;
      else if (ch === "[") bracketDepth += 1;
      else if (ch === "]") bracketDepth -= 1;

      if (depth !== 0 || bracketDepth !== 0) continue;

      if (ch === "?") {
        if (qIndex === -1) qIndex = i;
        else nested += 1;
        continue;
      }
      if (ch === ":" && qIndex !== -1) {
        if (nested === 0) return { q: qIndex, colon: i };
        nested -= 1;
      }
    }

    return null;
  }

  /**
   * @param {string} input
   * @returns {string}
   */
  function ternaryToIf(input) {
    const compact = stripOuterParens(input);
    const match = findTopLevelTernary(compact);
    if (!match) return compact.trim();

    const test = ternaryToIf(compact.slice(0, match.q));
    const consequent = ternaryToIf(compact.slice(match.q + 1, match.colon));
    const alternate = ternaryToIf(compact.slice(match.colon + 1));
    return `if ${test} then ${consequent} else ${alternate}`;
  }

  expr = ternaryToIf(expr);

  return (
    expr
      .replaceAll("&&", " and ")
      .replaceAll("||", " or ")
      // Replace != before !
      .replaceAll("!=", " <> ")
      .replaceAll("==", " = ")
      // Unary ! (best-effort; avoid touching " <> ")
      .replaceAll(/(^|[^\w])!(?!=)/g, "$1not ")
      .trim()
  );
}

/**
 * @param {QueryOperation} operation
 * @param {string} inputName
 * @returns {string}
 */
function operationToM(operation, inputName) {
  switch (operation.type) {
    case "selectColumns":
      return `Table.SelectColumns(${inputName}, ${valueToM(operation.columns)})`;
    case "removeColumns":
      return `Table.RemoveColumns(${inputName}, ${valueToM(operation.columns)})`;
    case "filterRows":
      return `Table.SelectRows(${inputName}, each ${predicateToM(operation.predicate)})`;
    case "sortRows": {
      const specs = operation.sortBy.map((s) => {
        const dir = s.direction === "descending" ? "Order.Descending" : "Order.Ascending";
        return `{${escapeMString(s.column)}, ${dir}}`;
      });
      return `Table.Sort(${inputName}, {${specs.join(", ")}})`;
    }
    case "groupBy": {
      const aggs = operation.aggregations.map((a) => [a.as ?? `${a.op} of ${a.column}`, `each ${aggregationOpToM(a.op)}([${a.column}])`]);
      // `each ...` is emitted as a string; this keeps output in our supported subset.
      const aggText = `{${aggs.map((a) => `{${escapeMString(a[0])}, ${a[1]}}`).join(", ")}}`;
      return `Table.Group(${inputName}, ${valueToM(operation.groupColumns)}, ${aggText})`;
    }
    case "addColumn": {
      const body = jsFormulaToM(operation.formula);
      return `Table.AddColumn(${inputName}, ${escapeMString(operation.name)}, each ${body})`;
    }
    case "distinctRows": {
      const cols = operation.columns && operation.columns.length > 0 ? `, ${valueToM(operation.columns)}` : "";
      return `Table.Distinct(${inputName}${cols})`;
    }
    case "removeRowsWithErrors": {
      const cols = operation.columns && operation.columns.length > 0 ? `, ${valueToM(operation.columns)}` : "";
      return `Table.RemoveRowsWithErrors(${inputName}${cols})`;
    }
    case "renameColumn":
      return `Table.RenameColumns(${inputName}, {{${escapeMString(operation.oldName)}, ${escapeMString(operation.newName)}}})`;
    case "changeType": {
      return `Table.TransformColumnTypes(${inputName}, {{${escapeMString(operation.column)}, ${dataTypeToMTypeExpr(operation.newType)}}})`;
    }
    case "transformColumns": {
      const specs = operation.transforms.map((t) => {
        const body = jsFormulaToM(t.formula);
        const type = t.newType ? `, ${dataTypeToMTypeExpr(t.newType)}` : "";
        return `{${escapeMString(t.column)}, each ${body}${type}}`;
      });
      return `Table.TransformColumns(${inputName}, {${specs.join(", ")}})`;
    }
    case "pivot":
      return `Table.Pivot(${inputName}, {}, ${escapeMString(operation.rowColumn)}, ${escapeMString(operation.valueColumn)}, ${aggregationOpToM(operation.aggregation)})`;
    case "unpivot":
      return `Table.Unpivot(${inputName}, ${valueToM(operation.columns)}, ${escapeMString(operation.nameColumn)}, ${escapeMString(operation.valueColumn)})`;
    case "fillDown":
      return `Table.FillDown(${inputName}, ${valueToM(operation.columns)})`;
    case "replaceValues":
      return `Table.ReplaceValue(${inputName}, ${valueToM(operation.find)}, ${valueToM(operation.replace)}, Replacer.ReplaceValue, {${escapeMString(operation.column)}})`;
    case "splitColumn":
      return `Table.SplitColumn(${inputName}, ${escapeMString(operation.column)}, ${escapeMString(operation.delimiter)})`;
    default: {
      /** @type {never} */
      const exhausted = operation;
      throw new Error(`Unsupported operation '${exhausted.type}'`);
    }
  }
}

/**
 * Best-effort pretty-printer from the internal `Query` model back to an M
 * script in the supported subset.
 *
 * @param {Query} query
 * @returns {string}
 */
export function prettyPrintQueryToM(query) {
  const indent = "  ";
  const lines = ["let"];

  const bindings = [];
  bindings.push({ name: "Source", expr: sourceToM(query.source) });

  let prev = "Source";
  for (const step of query.steps) {
    const name = step.name || `Step ${bindings.length}`;
    bindings.push({ name, expr: operationToM(step.operation, toMIdentifier(prev)) });
    prev = name;
  }

  bindings.forEach((b, idx) => {
    const comma = idx === bindings.length - 1 ? "" : ",";
    lines.push(`${indent}${toMIdentifier(b.name)} = ${b.expr}${comma}`);
  });

  lines.push("in");
  lines.push(`${indent}${toMIdentifier(prev)}`);
  return lines.join("\n");
}
