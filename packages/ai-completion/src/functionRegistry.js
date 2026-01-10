/**
 * @typedef {"range"|"value"|"number"|"string"|"boolean"|"any"} ArgType
 *
 * For tab-completion we only need a coarse idea of argument intent.
 * Excel's real signatures are significantly more nuanced (e.g. SUM accepts
 * both numbers and ranges, VLOOKUP's 4th arg is optional, etc.).
 */

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
 *   args: FunctionArgSpec[]
 * }} FunctionSpec
 */

export class FunctionRegistry {
  /**
   * @param {FunctionSpec[]} [functions]
   */
  constructor(functions) {
    /** @type {Map<string, FunctionSpec>} */
    this.functionsByName = new Map();
    for (const fn of functions ?? DEFAULT_FUNCTIONS) {
      this.register(fn);
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
    return this.functionsByName.get(name.toUpperCase());
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
    const matches = [];
    for (const spec of this.functionsByName.values()) {
      if (spec.name.startsWith(normalized)) matches.push(spec);
    }
    matches.sort((a, b) => a.name.localeCompare(b.name));
    return matches.slice(0, limit);
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
const DEFAULT_FUNCTIONS = [
  {
    name: "SUM",
    description: "Adds all the numbers in a range of cells.",
    args: [
      { name: "number1", type: "range", repeating: true },
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
];
