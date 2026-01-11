/**
 * Minimal M "standard library" surface for compilation into the power-query
 * `Query` model.
 */

/**
 * @typedef {import("../model.js").DataType} DataType
 */

export const TABLE_FUNCTION_STEP_NAMES = {
  "Table.SelectColumns": "Selected Columns",
  "Table.RemoveColumns": "Removed Columns",
  "Table.Sort": "Sorted Rows",
  "Table.SelectRows": "Filtered Rows",
  "Table.Distinct": "Removed Duplicates",
  "Table.Group": "Grouped Rows",
  "Table.AddColumn": "Added Column",
  "Table.RenameColumns": "Renamed Columns",
  "Table.TransformColumnTypes": "Changed Type",
  "Table.TransformColumns": "Transformed Columns",
  "Table.Pivot": "Pivoted Column",
  "Table.Unpivot": "Unpivoted Columns",
  "Table.Join": "Merged Queries",
  "Table.NestedJoin": "Merged Queries",
  "Table.Combine": "Appended Queries",
  "Table.RemoveRowsWithErrors": "Removed Errors",
  "Table.FillDown": "Filled Down",
  "Table.ReplaceValue": "Replaced Value",
  "Table.SplitColumn": "Split Column",
};

export const TABLE_FUNCTIONS = new Set(Object.keys(TABLE_FUNCTION_STEP_NAMES));

export const SOURCE_FUNCTIONS = new Set([
  "Excel.CurrentWorkbook",
  "Csv.Document",
  "Json.Document",
  "Web.Contents",
  "OData.Feed",
  "Odbc.Query",
  "Sql.Database",
  "Range.FromValues",
  "Query.Reference",
]);

/**
 * @param {string} fnName
 * @returns {string}
 */
export function defaultStepName(fnName) {
  return TABLE_FUNCTION_STEP_NAMES[fnName] ?? fnName;
}

/**
 * @param {string[]} parts
 * @returns {string}
 */
export function identifierPartsToName(parts) {
  return parts.join(".");
}

/**
 * @param {string} name
 * @returns {DataType}
 */
export function mTypeNameToDataType(name) {
  const lower = name.trim().toLowerCase();
  if (lower === "number") return "number";
  if (lower === "text" || lower === "string") return "string";
  if (lower === "logical" || lower === "bool" || lower === "boolean") return "boolean";
  if (lower === "date") return "date";
  return "any";
}

/**
 * M often uses `Int64.Type`, `Text.Type`, etc.
 * @param {string} name
 * @returns {DataType | null}
 */
export function identifierToDataType(name) {
  const lower = name.toLowerCase();
  if (lower === "int64.type" || lower === "double.type" || lower === "currency.type" || lower === "number.type") {
    return "number";
  }
  if (lower === "text.type" || lower === "string.type") return "string";
  if (lower === "logical.type" || lower === "bool.type" || lower === "boolean.type") return "boolean";
  if (lower === "date.type") return "date";
  if (lower === "any.type") return "any";
  return null;
}

/**
 * @param {string} name
 * @returns {unknown}
 */
export function constantIdentifierValue(name) {
  switch (name) {
    case "Order.Ascending":
      return "ascending";
    case "Order.Descending":
      return "descending";
    case "Nulls.First":
      return "first";
    case "Nulls.Last":
      return "last";
    case "Comparer.Ordinal":
      return { comparer: "ordinal", caseSensitive: true };
    case "Comparer.OrdinalIgnoreCase":
      return { comparer: "ordinalIgnoreCase", caseSensitive: false };
    default:
      return undefined;
  }
}

/**
 * @param {string} name
 * @returns {"sum" | "count" | "average" | "min" | "max" | "countDistinct" | null}
 */
export function listAggregationFromIdentifier(name) {
  switch (name) {
    case "List.Sum":
      return "sum";
    case "List.Count":
      return "count";
    case "List.Average":
      return "average";
    case "List.Min":
      return "min";
    case "List.Max":
      return "max";
    case "List.CountDistinct":
      return "countDistinct";
    default:
      return null;
  }
}
