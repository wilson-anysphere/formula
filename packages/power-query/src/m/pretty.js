/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QuerySource} QuerySource
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").FilterPredicate} FilterPredicate
 * @typedef {import("../model.js").ComparisonPredicate} ComparisonPredicate
 * @typedef {import("../model.js").Aggregation} Aggregation
 */

import { parseFormula } from "../expr/index.js";
import { MS_PER_DAY, PqDateTimeZone, PqDecimal, PqDuration, PqTime, hasUtcTimeComponent } from "../values.js";

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
  if (value instanceof PqDecimal) {
    return `Decimal.FromText(${escapeMString(value.value)})`;
  }
  if (value instanceof PqDateTimeZone) {
    const local = new Date(value.date.getTime() + value.offsetMinutes * 60 * 1000);
    const seconds = local.getUTCSeconds();
    const millis = local.getUTCMilliseconds();
    const secArg = millis === 0 ? String(seconds) : `${seconds}.${String(millis).padStart(3, "0")}`;

    const sign = value.offsetMinutes < 0 ? -1 : 1;
    const abs = Math.abs(value.offsetMinutes);
    const offH = sign * Math.floor(abs / 60);
    const offM = sign * (abs % 60);

    return `#datetimezone(${local.getUTCFullYear()}, ${local.getUTCMonth() + 1}, ${local.getUTCDate()}, ${local.getUTCHours()}, ${local.getUTCMinutes()}, ${secArg}, ${offH}, ${offM})`;
  }
  if (value instanceof PqTime) {
    const ms = value.milliseconds;
    const hours = Math.floor(ms / 3_600_000);
    const minutes = Math.floor((ms % 3_600_000) / 60_000);
    const seconds = Math.floor((ms % 60_000) / 1000);
    const millis = Math.floor(ms % 1000);
    const secArg = millis === 0 ? String(seconds) : `${seconds}.${String(millis).padStart(3, "0")}`;
    return `#time(${hours}, ${minutes}, ${secArg})`;
  }
  if (value instanceof PqDuration) {
    const sign = value.milliseconds < 0 ? -1 : 1;
    let remaining = Math.abs(value.milliseconds);
    const days = Math.floor(remaining / MS_PER_DAY);
    remaining -= days * MS_PER_DAY;
    const hours = Math.floor(remaining / 3_600_000);
    remaining -= hours * 3_600_000;
    const minutes = Math.floor(remaining / 60_000);
    remaining -= minutes * 60_000;
    const seconds = Math.floor(remaining / 1000);
    const millis = Math.floor(remaining - seconds * 1000);
    const secArg = millis === 0 ? String(sign * seconds) : `${sign < 0 ? "-" : ""}${seconds}.${String(millis).padStart(3, "0")}`;
    return `#duration(${sign * days}, ${sign * hours}, ${sign * minutes}, ${secArg})`;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    if (!hasUtcTimeComponent(value)) {
      return `#date(${value.getUTCFullYear()}, ${value.getUTCMonth() + 1}, ${value.getUTCDate()})`;
    }
    const seconds = value.getUTCSeconds();
    const millis = value.getUTCMilliseconds();
    const secArg = millis === 0 ? String(seconds) : `${seconds}.${String(millis).padStart(3, "0")}`;
    return `#datetime(${value.getUTCFullYear()}, ${value.getUTCMonth() + 1}, ${value.getUTCDate()}, ${value.getUTCHours()}, ${value.getUTCMinutes()}, ${secArg})`;
  }
  if (value instanceof Uint8Array) {
    let base64;
    if (typeof Buffer !== "undefined") {
      base64 = Buffer.from(value).toString("base64");
    } else {
      let binary = "";
      for (let i = 0; i < value.length; i++) binary += String.fromCharCode(value[i]);
      // eslint-disable-next-line no-undef
      base64 = btoa(binary);
    }
    return `Binary.FromText(${escapeMString(base64)}, BinaryEncoding.Base64)`;
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
    case "decimal":
      return "Decimal.Type";
    case "boolean":
      return "type logical";
    case "date":
      return "type date";
    case "datetime":
      return "type datetime";
    case "datetimezone":
      return "type datetimezone";
    case "time":
      return "type time";
    case "duration":
      return "type duration";
    case "binary":
      return "type binary";
    case "any":
    default:
      return "type any";
  }
}

/**
 * @param {"inner" | "left" | "right" | "full"} joinType
 * @returns {string}
 */
function joinTypeToM(joinType) {
  switch (joinType) {
    case "inner":
      return "JoinKind.Inner";
    case "left":
      return "JoinKind.LeftOuter";
    case "right":
      return "JoinKind.RightOuter";
    case "full":
      return "JoinKind.FullOuter";
    default:
      return "JoinKind.Inner";
  }
}

/**
 * @param {{ comparer: string; caseSensitive?: boolean } | null | undefined} comparer
 * @returns {string}
 */
function joinComparerToM(comparer) {
  if (!comparer) return "Comparer.Ordinal";
  const name = typeof comparer.comparer === "string" ? comparer.comparer.toLowerCase() : "";
  if (comparer.caseSensitive === false || name === "ordinalignorecase") return "Comparer.OrdinalIgnoreCase";
  return "Comparer.Ordinal";
}

/**
 * @param {string} formula
 * @returns {string}
 */
function jsFormulaToM(formula) {
  try {
    const expr = parseFormula(formula);

    /**
     * @param {import("../expr/ast.js").ExprNode} node
     * @returns {string}
     */
    function toM(node) {
      switch (node.type) {
        case "literal":
          return valueToM(node.value);
        case "column":
          return `[${node.name}]`;
        case "value":
          return "_";
        case "unary": {
          const arg = toM(node.arg);
          if (node.op === "!") return `not (${arg})`;
          return `(${node.op}(${arg}))`;
        }
        case "binary": {
          const left = toM(node.left);
          const right = toM(node.right);
          const op = (() => {
            switch (node.op) {
              case "&&":
                return "and";
              case "||":
                return "or";
              case "==":
              case "===":
                return "=";
              case "!=":
              case "!==":
                return "<>";
              default:
                return node.op;
            }
          })();
          return `(${left} ${op} ${right})`;
        }
        case "ternary": {
          const test = toM(node.test);
          const consequent = toM(node.consequent);
          const alternate = toM(node.alternate);
          return `(if ${test} then ${consequent} else ${alternate})`;
        }
        case "call": {
          const callee = node.callee.toLowerCase();
          const name =
            callee === "text_upper"
              ? "Text.Upper"
              : callee === "text_lower"
                ? "Text.Lower"
                : callee === "text_trim"
                  ? "Text.Trim"
                  : callee === "text_length"
                    ? "Text.Length"
                    : callee === "text_contains"
                      ? "Text.Contains"
                      : callee === "number_round"
                        ? "Number.Round"
                        : callee === "date_add_days"
                          ? "Date.AddDays"
                          : callee === "date_from_text" || callee === "date"
                            ? "Date.FromText"
                            : node.callee;
          return `${name}(${node.args.map(toM).join(", ")})`;
        }
        default: {
          /** @type {never} */
          const exhausted = node;
          throw new Error(`Unsupported node '${exhausted.type}'`);
        }
      }
    }

    return toM(expr);
  } catch {
    // Fallback: return something parseable in our subset.
    let expr = formula.trim();
    if (expr.startsWith("=")) expr = expr.slice(1).trim();
    return expr;
  }
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
      return `Table.SplitColumn(${inputName}, ${escapeMString(operation.column)}, ${escapeMString(operation.delimiter)}${
        operation.newColumns && operation.newColumns.length > 0 ? `, ${valueToM(operation.newColumns)}` : ""
      })`;
    case "promoteHeaders":
      return `Table.PromoteHeaders(${inputName})`;
    case "demoteHeaders":
      return `Table.DemoteHeaders(${inputName})`;
    case "reorderColumns": {
      const missing =
        operation.missingField && operation.missingField !== "error"
          ? `, ${operation.missingField === "ignore" ? "MissingField.Ignore" : "MissingField.UseNull"}`
          : "";
      return `Table.ReorderColumns(${inputName}, ${valueToM(operation.columns)}${missing})`;
    }
    case "addIndexColumn": {
      const initial = Number.isFinite(operation.initialValue) ? operation.initialValue : 0;
      const increment = Number.isFinite(operation.increment) ? operation.increment : 1;
      const optional = initial === 0 && increment === 1 ? "" : `, ${valueToM(initial)}, ${valueToM(increment)}`;
      return `Table.AddIndexColumn(${inputName}, ${escapeMString(operation.name)}${optional})`;
    }
    case "take":
      return `Table.FirstN(${inputName}, ${valueToM(operation.count)})`;
    case "skip":
      return `Table.Skip(${inputName}, ${valueToM(operation.count)})`;
    case "removeRows":
      return `Table.RemoveRows(${inputName}, ${valueToM(operation.offset)}, ${valueToM(operation.count)})`;
    case "combineColumns":
      return `Table.CombineColumns(${inputName}, ${valueToM(operation.columns)}, Combiner.CombineTextByDelimiter(${escapeMString(operation.delimiter)}, QuoteStyle.None), ${escapeMString(operation.newColumnName)})`;
    case "transformColumnNames": {
      const fn = operation.transform === "upper" ? "Text.Upper" : operation.transform === "lower" ? "Text.Lower" : "Text.Trim";
      return `Table.TransformColumnNames(${inputName}, ${fn})`;
    }
    case "replaceErrorValues": {
      const specs = operation.replacements.map((r) => `{${escapeMString(r.column)}, ${valueToM(r.value)}}`);
      return `Table.ReplaceErrorValues(${inputName}, {${specs.join(", ")}})`;
    }
    case "append": {
      const others = operation.queries.map((id) => `Query.Reference(${escapeMString(id)})`);
      return `Table.Combine({${[inputName, ...others].join(", ")}})`;
    }
    case "merge": {
      const leftKeys =
        Array.isArray(operation.leftKeys) && operation.leftKeys.length > 0
          ? operation.leftKeys
          : typeof operation.leftKey === "string" && operation.leftKey
            ? [operation.leftKey]
            : [];
      const rightKeys =
        Array.isArray(operation.rightKeys) && operation.rightKeys.length > 0
          ? operation.rightKeys
          : typeof operation.rightKey === "string" && operation.rightKey
            ? [operation.rightKey]
            : [];
      const joinMode = operation.joinMode ?? "flat";
      const joinKind = joinTypeToM(operation.joinType);
      const right = `Query.Reference(${escapeMString(operation.rightQuery)})`;
      const comparerArg =
        operation.comparer != null ? `, null, ${joinComparerToM(operation.comparer)}` : "";
      if (joinMode === "nested") {
        if (typeof operation.newColumnName !== "string") {
          throw new Error("Nested join requires newColumnName");
        }
        return `Table.NestedJoin(${inputName}, ${valueToM(leftKeys)}, ${right}, ${valueToM(rightKeys)}, ${escapeMString(operation.newColumnName)}, ${joinKind}${comparerArg})`;
      }
      return `Table.Join(${inputName}, ${valueToM(leftKeys)}, ${right}, ${valueToM(rightKeys)}, ${joinKind}${comparerArg})`;
    }
    case "expandTableColumn": {
      const cols = operation.columns == null ? "null" : valueToM(operation.columns);
      const names = operation.newColumnNames == null ? "" : `, ${valueToM(operation.newColumnNames)}`;
      return `Table.ExpandTableColumn(${inputName}, ${escapeMString(operation.column)}, ${cols}${names})`;
    }
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
