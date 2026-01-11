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
    const spec = fn.args[argIndex] ?? fn.args.find(a => a.repeating);
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
    name: "AVERAGE",
    description: "Returns the average (arithmetic mean) of the arguments.",
    args: [
      { name: "number1", type: "range", repeating: true },
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
    name: "MIN",
    description: "Returns the smallest value in a set of values.",
    args: [
      { name: "number1", type: "range", repeating: true },
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
    name: "TRANSPOSE",
    description: "Returns the transpose of an array or range.",
    args: [
      { name: "array", type: "range" },
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
    out.push({
      name: entry.name.toUpperCase(),
      minArgs,
      maxArgs,
      args: [],
    });
  }
  return out;
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
