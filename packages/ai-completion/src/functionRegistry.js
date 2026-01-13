/**
 * @typedef {"range"|"value"|"number"|"string"|"boolean"|"any"} ArgType
 *
 * For tab-completion we only need a coarse idea of argument intent.
 * Excel's real signatures are significantly more nuanced (e.g. SUM accepts
 * both numbers and ranges, VLOOKUP's 4th arg is optional, etc.).
 */

import FUNCTION_CATALOG from "../../../shared/functionCatalog.mjs";

/**
 * @typedef {{
 *   name: string,
 *   type: ArgType,
 *   optional?: boolean,
 *   repeating?: boolean
 * }} FunctionArgSpec
 */

/**
 * @typedef {{
 *   name: string,
 *   description?: string,
 *   minArgs?: number,
 *   maxArgs?: number,
 *   args: FunctionArgSpec[]
 * }} FunctionSpec
 */

export class FunctionRegistry {
  /**
   * @param {FunctionSpec[]} [functions]
   * @param {{ catalog?: any }} [options]
   */
  constructor(functions, options = {}) {
    /** @type {Map<string, FunctionSpec>} */
    this.functionsByName = new Map();

    /** @type {string[] | null} */
    this._sortedUpperNames = null;

    if (Array.isArray(functions)) {
      for (const fn of functions) this.register(fn);
      return;
    }

    const catalogSource = Object.prototype.hasOwnProperty.call(options, "catalog")
      ? options.catalog
      : FUNCTION_CATALOG;
    const catalogFunctions = functionsFromCatalog(catalogSource);

    for (const fn of catalogFunctions) {
      this.register(fn);
      const xlfnAlias = toXlfnAlias(fn);
      if (xlfnAlias) this.register(xlfnAlias);
    }

    // Signature/arg-type hints are intentionally curated until the catalog
    // carries richer metadata. Curated entries override catalog name-only
    // entries.
    for (const fn of CURATED_FUNCTIONS) {
      const key = fn.name.toUpperCase();
      const existing = this.functionsByName.get(key);
      const merged = existing ? { ...existing, ...fn } : fn;
      this.register(merged);
      const xlfnAlias = toXlfnAlias(merged);
      if (xlfnAlias) this.register(xlfnAlias);
    }
  }

  /**
   * @param {FunctionSpec} spec
   */
  register(spec) {
    if (!spec || typeof spec.name !== "string" || spec.name.length === 0) {
      throw new Error("FunctionRegistry.register: function name is required");
    }
    this.functionsByName.set(spec.name.toUpperCase(), spec);
    this._sortedUpperNames = null;
  }

  /**
   * @returns {FunctionSpec[]}
   */
  list() {
    return [...this.functionsByName.values()];
  }

  /**
   * @param {string} name
   * @returns {FunctionSpec | undefined}
   */
  getFunction(name) {
    if (typeof name !== "string") return undefined;
    const upper = name.toUpperCase();
    const direct = this.functionsByName.get(upper);
    if (direct) return direct;
    if (upper.startsWith("_XLFN.")) {
      return this.functionsByName.get(upper.slice("_XLFN.".length));
    }
    return undefined;
  }

  /**
   * @param {string} prefix
   * @param {{limit?: number}} [options]
   * @returns {FunctionSpec[]}
   */
  search(prefix, options = {}) {
    const limit = options.limit ?? 10;
    if (typeof prefix !== "string" || prefix.length === 0) return [];
    const normalized = prefix.toUpperCase();

    if (!this._sortedUpperNames) {
      this._sortedUpperNames = [...this.functionsByName.keys()].sort();
    }

    const names = this._sortedUpperNames;
    const start = lowerBound(names, normalized);

    /** @type {FunctionSpec[]} */
    const matches = [];
    for (let i = start; i < names.length && matches.length < limit; i++) {
      const name = names[i];
      if (!name.startsWith(normalized)) break;
      const spec = this.functionsByName.get(name);
      if (spec) matches.push(spec);
    }

    return matches;
  }

  /**
   * @param {string} functionName
   * @param {number} argIndex 0-based
   * @returns {ArgType | undefined}
   */
  getArgType(functionName, argIndex) {
    const fn = this.getFunction(functionName);
    if (!fn) return undefined;
    if (!Number.isInteger(argIndex) || argIndex < 0) return undefined;
    if (!Array.isArray(fn.args) || fn.args.length === 0) return undefined;

    const direct = fn.args[argIndex];
    if (direct) return direct.type;

    // Support repeating "groups" of arguments (e.g. SUMIFS(range1, crit1, range2, crit2, ...)).
    // The first arg with `repeating:true` marks the start of a group that repeats to `maxArgs`.
    const repeatingStart = fn.args.findIndex((a) => a.repeating);
    if (repeatingStart < 0) return undefined;

    const groupLen = fn.args.length - repeatingStart;
    if (groupLen <= 0) return undefined;

    const offset = (argIndex - repeatingStart) % groupLen;
    const spec = fn.args[repeatingStart + offset];
    return spec?.type;
  }

  /**
   * @param {string} functionName
   * @param {number} argIndex
   */
  isRangeArg(functionName, argIndex) {
    const type = this.getArgType(functionName, argIndex);
    return type === "range";
  }
}

/** @type {FunctionSpec[]} */
const CURATED_FUNCTIONS = [
  {
    name: "SUM",
    description: "Adds all the numbers in a range of cells.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "SUMIF",
    description: "Adds the cells specified by a given condition or criteria.",
    args: [
      { name: "range", type: "range" },
      { name: "criteria", type: "value" },
      { name: "sum_range", type: "range", optional: true },
    ],
  },
  {
    name: "SUMIFS",
    description: "Adds the cells in a range that meet multiple criteria.",
    args: [
      { name: "sum_range", type: "range" },
      { name: "criteria_range1", type: "range", repeating: true },
      { name: "criteria1", type: "value" },
    ],
  },
  {
    name: "COUNT",
    description: "Counts the number of cells that contain numbers.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "COUNTA",
    description: "Counts the number of non-empty cells.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "COUNTBLANK",
    description: "Counts the number of blank cells within a range.",
    args: [
      { name: "range", type: "range", repeating: true },
    ],
  },
  {
    name: "COUNTIF",
    description: "Counts the number of cells within a range that meet the given criteria.",
    args: [
      { name: "range", type: "range" },
      { name: "criteria", type: "value" },
    ],
  },
  {
    name: "COUNTIFS",
    description: "Counts the number of cells that meet multiple criteria.",
    args: [
      { name: "criteria_range1", type: "range", repeating: true },
      { name: "criteria1", type: "value" },
    ],
  },
  {
    name: "AVERAGE",
    description: "Returns the average (arithmetic mean) of the arguments.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "AVERAGEIF",
    description: "Returns the average of cells that meet a given condition or criteria.",
    args: [
      { name: "range", type: "range" },
      { name: "criteria", type: "value" },
      { name: "average_range", type: "range", optional: true },
    ],
  },
  {
    name: "AVERAGEIFS",
    description: "Returns the average of cells that meet multiple criteria.",
    args: [
      { name: "average_range", type: "range" },
      { name: "criteria_range1", type: "range", repeating: true },
      { name: "criteria1", type: "value" },
    ],
  },
  {
    name: "MAX",
    description: "Returns the largest value in a set of values.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "MAXIFS",
    description: "Returns the maximum value among cells specified by a given set of conditions or criteria.",
    args: [
      { name: "max_range", type: "range" },
      { name: "criteria_range1", type: "range", repeating: true },
      { name: "criteria1", type: "value" },
    ],
  },
  {
    name: "MIN",
    description: "Returns the smallest value in a set of values.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "MINIFS",
    description: "Returns the minimum value among cells specified by a given set of conditions or criteria.",
    args: [
      { name: "min_range", type: "range" },
      { name: "criteria_range1", type: "range", repeating: true },
      { name: "criteria1", type: "value" },
    ],
  },
  {
    name: "SUMPRODUCT",
    description: "Returns the sum of the products of corresponding array components.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range", repeating: true },
    ],
  },
  {
    name: "LOOKUP",
    description: "Looks up a value either from a one-row or one-column range or from an array.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "lookup_vector", type: "range" },
      { name: "result_vector", type: "range", optional: true },
    ],
  },
  {
    name: "VLOOKUP",
    description: "Looks for a value in the leftmost column of a table.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "table_array", type: "range" },
      { name: "col_index_num", type: "number" },
      { name: "range_lookup", type: "boolean", optional: true },
    ],
  },
  {
    name: "HLOOKUP",
    description: "Looks for a value in the top row of a table.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "table_array", type: "range" },
      { name: "row_index_num", type: "number" },
      { name: "range_lookup", type: "boolean", optional: true },
    ],
  },
  {
    name: "XLOOKUP",
    description: "Looks up a value in a range or an array.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "lookup_array", type: "range" },
      { name: "return_array", type: "range" },
      { name: "if_not_found", type: "value", optional: true },
      { name: "match_mode", type: "number", optional: true },
      { name: "search_mode", type: "number", optional: true },
    ],
  },
  {
    name: "NPV",
    description: "Returns the net present value of an investment based on a series of periodic cash flows and a discount rate.",
    args: [
      { name: "rate", type: "number" },
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "IRR",
    description: "Returns the internal rate of return for a series of cash flows.",
    args: [
      { name: "values", type: "range" },
      { name: "guess", type: "number", optional: true },
    ],
  },
  {
    name: "XNPV",
    description: "Returns the net present value for a schedule of cash flows that is not necessarily periodic.",
    args: [
      { name: "rate", type: "number" },
      { name: "values", type: "range" },
      { name: "dates", type: "range" },
    ],
  },
  {
    name: "XIRR",
    description: "Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.",
    args: [
      { name: "values", type: "range" },
      { name: "dates", type: "range" },
      { name: "guess", type: "number", optional: true },
    ],
  },
  {
    name: "INDEX",
    description: "Returns the value of an element in a table or an array.",
    args: [
      { name: "array", type: "range" },
      { name: "row_num", type: "number" },
      { name: "column_num", type: "number", optional: true },
    ],
  },
  {
    name: "MATCH",
    description: "Looks up values in a reference or array.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "lookup_array", type: "range" },
      { name: "match_type", type: "number", optional: true },
    ],
  },
  {
    name: "XMATCH",
    description: "Returns the relative position of an item in an array or range.",
    args: [
      { name: "lookup_value", type: "value" },
      { name: "lookup_array", type: "range" },
      { name: "match_mode", type: "number", optional: true },
      { name: "search_mode", type: "number", optional: true },
    ],
  },
  {
    name: "IF",
    description: "Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.",
    args: [
      { name: "logical_test", type: "value" },
      { name: "value_if_true", type: "value" },
      { name: "value_if_false", type: "value", optional: true },
    ],
  },
  {
    name: "CONCAT",
    description: "Combines the text from multiple ranges and/or strings.",
    args: [
      { name: "text1", type: "value", repeating: true },
    ],
  },
  {
    name: "CONCATENATE",
    description: "Combines several text strings into one text string.",
    args: [
      { name: "text1", type: "value", repeating: true },
    ],
  },
  {
    name: "TEXTJOIN",
    description: "Combines text from multiple ranges and/or strings, and includes a delimiter between each value.",
    args: [
      { name: "delimiter", type: "string" },
      { name: "ignore_empty", type: "boolean" },
      { name: "text1", type: "range", repeating: true },
    ],
  },
  {
    name: "TRANSPOSE",
    description: "Returns the transpose of an array or range.",
    args: [
      { name: "array", type: "range" },
    ],
  },
  {
    name: "FILTER",
    description: "Filters a range of data based on criteria you define.",
    args: [
      { name: "array", type: "range" },
      { name: "include", type: "range" },
      { name: "if_empty", type: "value", optional: true },
    ],
  },
  {
    name: "SORT",
    description: "Sorts the contents of a range or array.",
    args: [
      { name: "array", type: "range" },
      { name: "sort_index", type: "number", optional: true },
      { name: "sort_order", type: "number", optional: true },
      { name: "by_col", type: "boolean", optional: true },
    ],
  },
  {
    name: "SORTBY",
    description: "Sorts the contents of a range or array based on the values in a corresponding range or array.",
    args: [
      { name: "array", type: "range" },
      { name: "by_array1", type: "range", repeating: true },
      { name: "sort_order1", type: "number", optional: true },
    ],
  },
  {
    name: "UNIQUE",
    description: "Returns a list of unique values in a list or range.",
    args: [
      { name: "array", type: "range" },
      { name: "by_col", type: "boolean", optional: true },
      { name: "exactly_once", type: "boolean", optional: true },
    ],
  },
  {
    name: "TAKE",
    description: "Returns a specified number of rows or columns from the start or end of an array.",
    args: [
      { name: "array", type: "range" },
      { name: "rows", type: "number", optional: true },
      { name: "columns", type: "number", optional: true },
    ],
  },
  {
    name: "DROP",
    description: "Excludes a specified number of rows or columns from the start or end of an array.",
    args: [
      { name: "array", type: "range" },
      { name: "rows", type: "number", optional: true },
      { name: "columns", type: "number", optional: true },
    ],
  },
  {
    name: "CHOOSECOLS",
    description: "Returns the specified columns from an array.",
    args: [
      { name: "array", type: "range" },
      { name: "col_num", type: "number", repeating: true },
    ],
  },
  {
    name: "CHOOSEROWS",
    description: "Returns the specified rows from an array.",
    args: [
      { name: "array", type: "range" },
      { name: "row_num", type: "number", repeating: true },
    ],
  },
  {
    name: "EXPAND",
    description: "Expands an array to the specified row and column dimensions.",
    args: [
      { name: "array", type: "range" },
      { name: "rows", type: "number" },
      { name: "columns", type: "number", optional: true },
      { name: "pad_with", type: "value", optional: true },
    ],
  },
  {
    name: "TEXTSPLIT",
    description: "Splits text into rows and columns using delimiters.",
    minArgs: 2,
    maxArgs: 6,
    args: [
      { name: "text", type: "range" },
      { name: "col_delimiter", type: "string" },
      { name: "row_delimiter", type: "string", optional: true },
      { name: "ignore_empty", type: "boolean", optional: true },
      { name: "match_mode", type: "number", optional: true },
      { name: "pad_with", type: "value", optional: true },
    ],
  },
];

/**
 * Turn the Rust-generated catalog shape into name-only FunctionSpec entries.
 * Validation is intentionally lenient so developer environments keep working
 * even if the catalog is missing or corrupted.
 *
 * @param {any} catalog
 * @returns {FunctionSpec[]}
 */
function functionsFromCatalog(catalog) {
  if (!catalog || typeof catalog !== "object" || !Array.isArray(catalog.functions)) {
    return [];
  }

  /** @type {FunctionSpec[]} */
  const out = [];
  for (const entry of catalog.functions) {
    if (!entry || typeof entry.name !== "string" || entry.name.length === 0) continue;
    const minArgs = Number.isInteger(entry.min_args) ? entry.min_args : undefined;
    const maxArgs = Number.isInteger(entry.max_args) ? entry.max_args : undefined;

    /** @type {FunctionArgSpec[]} */
    const args = [];
    const catalogArgTypes = Array.isArray(entry.arg_types) ? entry.arg_types : [];
    for (let i = 0; i < catalogArgTypes.length; i++) {
      const type = catalogValueTypeToArgType(catalogArgTypes[i]);
      if (!type) continue;
      args.push({
        name: `arg${i + 1}`,
        type,
        optional: Number.isInteger(minArgs) ? i >= minArgs : undefined,
      });
    }

    if (Number.isInteger(maxArgs) && args.length > 0 && maxArgs > args.length) {
      // Best-effort varargs handling: many Excel functions define a single ValueType
      // and allow 0..255 args. We treat the last typed arg as repeating so argIndex
      // lookups don’t fall off a cliff.
      args[args.length - 1].repeating = true;
    }
    out.push({
      name: entry.name.toUpperCase(),
      minArgs,
      maxArgs,
      args,
    });
  }
  return out;
}

/**
 * @param {unknown} valueType
 * @returns {ArgType | null}
 */
function catalogValueTypeToArgType(valueType) {
  switch (valueType) {
    case "number":
      return "number";
    case "bool":
      return "boolean";
    case "text":
    case "any":
      return "value";
    default:
      return null;
  }
}

/**
 * @param {string[]} arr Sorted array.
 * @param {string} target
 */
function lowerBound(arr, target) {
  let lo = 0;
  let hi = arr.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (arr[mid] < target) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * Excel stores some "newer" functions in formula text with an `_xlfn.` prefix.
 * For completion/hinting we treat these as aliases of the unprefixed function.
 *
 * @param {FunctionSpec} spec
 * @returns {FunctionSpec | null}
 */
function toXlfnAlias(spec) {
  const baseName = spec?.name?.toUpperCase?.();
  if (!baseName || typeof baseName !== "string") return null;
  if (baseName.startsWith("_XLFN.")) return null;
  return { ...spec, name: `_XLFN.${baseName}` };
}
