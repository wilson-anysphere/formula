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
    name: "SUBTOTAL",
    description: "Returns a subtotal in a list or database.",
    args: [
      { name: "function_num", type: "number" },
      { name: "ref1", type: "range", repeating: true },
    ],
  },
  {
    name: "AGGREGATE",
    description: "Returns an aggregate in a list or database, ignoring hidden rows and errors.",
    args: [
      { name: "function_num", type: "number" },
      { name: "options", type: "number" },
      { name: "ref1", type: "range", repeating: true },
    ],
  },
  {
    name: "WORKDAY",
    description: "Returns a date that is a specified number of working days before or after a start date.",
    args: [
      { name: "start_date", type: "number" },
      { name: "days", type: "number" },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "DATEDIF",
    description: "Calculates the number of days, months, or years between two dates.",
    args: [
      { name: "start_date", type: "value" },
      { name: "end_date", type: "value" },
      { name: "unit", type: "string" },
    ],
  },
  {
    name: "DAYS",
    description: "Returns the number of days between two dates.",
    args: [
      { name: "end_date", type: "value" },
      { name: "start_date", type: "value" },
    ],
  },
  {
    name: "DAYS360",
    description: "Returns the number of days between two dates based on a 360-day year.",
    args: [
      { name: "start_date", type: "value" },
      { name: "end_date", type: "value" },
      { name: "method", type: "boolean", optional: true },
    ],
  },
  {
    name: "YEARFRAC",
    description: "Returns the year fraction representing the number of whole days between start_date and end_date.",
    args: [
      { name: "start_date", type: "value" },
      { name: "end_date", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "DATEVALUE",
    description: "Converts a date in the form of text to a serial number.",
    args: [
      { name: "date_text", type: "value" },
    ],
  },
  {
    name: "TIMEVALUE",
    description: "Converts a time in the form of text to a serial number.",
    args: [
      { name: "time_text", type: "value" },
    ],
  },
  {
    name: "HOUR",
    description: "Converts a serial number to an hour.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "MINUTE",
    description: "Converts a serial number to a minute.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "SECOND",
    description: "Converts a serial number to a second.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "WORKDAY.INTL",
    description: "Returns a date that is a specified number of working days before or after a start date with custom weekends.",
    args: [
      { name: "start_date", type: "number" },
      { name: "days", type: "number" },
      { name: "weekend", type: "number", optional: true },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "NETWORKDAYS",
    description: "Returns the number of whole working days between two dates.",
    args: [
      { name: "start_date", type: "number" },
      { name: "end_date", type: "number" },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "NETWORKDAYS.INTL",
    description: "Returns the number of whole working days between two dates with custom weekends.",
    args: [
      { name: "start_date", type: "number" },
      { name: "end_date", type: "number" },
      { name: "weekend", type: "number", optional: true },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "BYROW",
    description: "Applies a LAMBDA to each row in an array and returns the result as an array.",
    args: [
      { name: "array", type: "range" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "BYCOL",
    description: "Applies a LAMBDA to each column in an array and returns the result as an array.",
    args: [
      { name: "array", type: "range" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "MAKEARRAY",
    description: "Returns an array of the specified size by applying a LAMBDA to each element.",
    args: [
      { name: "rows", type: "number" },
      { name: "columns", type: "number" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "DAVERAGE",
    description: "Averages values in a list or database that match conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DCOUNT",
    description: "Counts the cells that contain numbers in a list or database that match conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DCOUNTA",
    description: "Counts nonblank cells in a list or database that match conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DGET",
    description: "Extracts a single value from a list or database that matches conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DMAX",
    description: "Returns the largest number in a list or database that matches conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DMIN",
    description: "Returns the smallest number in a list or database that matches conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DPRODUCT",
    description: "Multiplies values in a list or database that match conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DSTDEV",
    description: "Estimates standard deviation based on a sample from selected database entries.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DSTDEVP",
    description: "Calculates standard deviation based on the entire population of selected database entries.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DSUM",
    description: "Adds numbers in a list or database that match conditions you specify.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DVAR",
    description: "Estimates variance based on a sample from selected database entries.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
    ],
  },
  {
    name: "DVARP",
    description: "Calculates variance based on the entire population of selected database entries.",
    args: [
      { name: "database", type: "range" },
      { name: "field", type: "value" },
      { name: "criteria", type: "range" },
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
    name: "AVERAGEA",
    description: "Returns the average of its arguments, including numbers, text, and logical values.",
    args: [
      { name: "value1", type: "range", repeating: true },
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
    name: "MEDIAN",
    description: "Returns the median of the given numbers.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEV.S",
    description: "Estimates standard deviation based on a sample.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEV.P",
    description: "Calculates standard deviation based on the entire population.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEVA",
    description: "Estimates standard deviation based on a sample, including logical values and text.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEVP",
    description: "Calculates standard deviation based on the entire population (legacy).",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEV",
    description: "Estimates standard deviation based on a sample (legacy).",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "STDEVPA",
    description: "Calculates standard deviation based on the entire population, including logical values and text.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "VAR.S",
    description: "Estimates variance based on a sample.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "VAR.P",
    description: "Calculates variance based on the entire population.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "VARA",
    description: "Estimates variance based on a sample, including logical values and text.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "VARP",
    description: "Calculates variance based on the entire population (legacy).",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "VARPA",
    description: "Calculates variance based on the entire population, including logical values and text.",
    args: [
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "VAR",
    description: "Estimates variance based on a sample (legacy).",
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
    name: "MAXA",
    description: "Returns the largest value in a list of arguments, including numbers, text, and logical values.",
    args: [
      { name: "value1", type: "range", repeating: true },
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
    name: "MINA",
    description: "Returns the smallest value in a list of arguments, including numbers, text, and logical values.",
    args: [
      { name: "value1", type: "range", repeating: true },
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
    name: "LARGE",
    description: "Returns the k-th largest value in a data set.",
    args: [
      { name: "array", type: "range" },
      { name: "k", type: "number" },
    ],
  },
  {
    name: "SMALL",
    description: "Returns the k-th smallest value in a data set.",
    args: [
      { name: "array", type: "range" },
      { name: "k", type: "number" },
    ],
  },
  {
    name: "PERCENTILE.INC",
    description: "Returns the k-th percentile of values in a range.",
    args: [
      { name: "array", type: "range" },
      { name: "k", type: "number" },
    ],
  },
  {
    name: "PERCENTILE.EXC",
    description: "Returns the k-th percentile of values in a range, exclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "k", type: "number" },
    ],
  },
  {
    name: "PERCENTILE",
    description: "Returns the k-th percentile of values in a range (legacy).",
    args: [
      { name: "array", type: "range" },
      { name: "k", type: "number" },
    ],
  },
  {
    name: "QUARTILE.INC",
    description: "Returns the quartile of a data set.",
    args: [
      { name: "array", type: "range" },
      { name: "quart", type: "number" },
    ],
  },
  {
    name: "QUARTILE.EXC",
    description: "Returns the quartile of a data set, exclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "quart", type: "number" },
    ],
  },
  {
    name: "QUARTILE",
    description: "Returns the quartile of a data set (legacy).",
    args: [
      { name: "array", type: "range" },
      { name: "quart", type: "number" },
    ],
  },
  {
    name: "RANK.EQ",
    description: "Returns the rank of a number in a list of numbers.",
    args: [
      { name: "number", type: "number" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "RANK",
    description: "Returns the rank of a number in a list of numbers (legacy).",
    args: [
      { name: "number", type: "number" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "RANK.AVG",
    description: "Returns the rank of a number in a list of numbers, with ties averaged.",
    args: [
      { name: "number", type: "number" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK",
    description: "Returns the rank of a value in a data set as a percentage of the data set.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "number" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK.INC",
    description: "Returns the rank of a value in a data set as a percentage of the data set, inclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "number" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK.EXC",
    description: "Returns the rank of a value in a data set as a percentage of the data set, exclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "number" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "MODE.SNGL",
    description: "Returns the most common value in a data set.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "MODE.MULT",
    description: "Returns a vertical array of the most frequently occurring, or repetitive values in an array or range of data.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "MODE",
    description: "Returns the most common value in a data set (legacy).",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "AVEDEV",
    description: "Returns the average of the absolute deviations of data points from their mean.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "DEVSQ",
    description: "Returns the sum of squares of deviations of data points from their sample mean.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "KURT",
    description: "Returns the kurtosis of a data set.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "SKEW",
    description: "Returns the skewness of a distribution.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "SKEW.P",
    description: "Returns the skewness of a population.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "GEOMEAN",
    description: "Returns the geometric mean.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "HARMEAN",
    description: "Returns the harmonic mean.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "TRIMMEAN",
    description: "Returns the mean of the interior portion of a data set.",
    args: [
      { name: "array", type: "range" },
      { name: "percent", type: "number" },
    ],
  },
  {
    name: "CORREL",
    description: "Returns the correlation coefficient between two data sets.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "COVAR",
    description: "Returns the covariance between two data sets.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "COVARIANCE.P",
    description: "Returns the population covariance between two data sets.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "COVARIANCE.S",
    description: "Returns the sample covariance between two data sets.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "PEARSON",
    description: "Returns the Pearson product moment correlation coefficient.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "CHISQ.TEST",
    description: "Returns the test for independence.",
    args: [
      { name: "actual_range", type: "range" },
      { name: "expected_range", type: "range" },
    ],
  },
  {
    name: "CHITEST",
    description: "Returns the test for independence (legacy).",
    args: [
      { name: "actual_range", type: "range" },
      { name: "expected_range", type: "range" },
    ],
  },
  {
    name: "F.TEST",
    description: "Returns the result of an F-test.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "FTEST",
    description: "Returns the result of an F-test (legacy).",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "T.TEST",
    description: "Returns the probability associated with a Student's t-test.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
      { name: "tails", type: "number" },
      { name: "type", type: "number" },
    ],
  },
  {
    name: "TTEST",
    description: "Returns the probability associated with a Student's t-test (legacy).",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
      { name: "tails", type: "number" },
      { name: "type", type: "number" },
    ],
  },
  {
    name: "Z.TEST",
    description: "Returns the one-tailed probability-value of a z-test.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "number" },
      { name: "sigma", type: "number", optional: true },
    ],
  },
  {
    name: "ZTEST",
    description: "Returns the one-tailed probability-value of a z-test (legacy).",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "number" },
      { name: "sigma", type: "number", optional: true },
    ],
  },
  {
    name: "FREQUENCY",
    description: "Calculates how often values occur within a range of values and returns a vertical array of numbers.",
    args: [
      { name: "data_array", type: "range" },
      { name: "bins_array", type: "range" },
    ],
  },
  {
    name: "PROB",
    description: "Returns the probability that values in a range are between two limits.",
    args: [
      { name: "x_range", type: "range" },
      { name: "prob_range", type: "range" },
      { name: "lower_limit", type: "number" },
      { name: "upper_limit", type: "number", optional: true },
    ],
  },
  {
    name: "SERIESSUM",
    description: "Returns the sum of a power series based on the formula.",
    args: [
      { name: "x", type: "number" },
      { name: "n", type: "number" },
      { name: "m", type: "number" },
      { name: "coefficients", type: "range" },
    ],
  },
  {
    name: "RSQ",
    description: "Returns the square of the Pearson product moment correlation coefficient.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "STEYX",
    description: "Returns the standard error of the predicted y-value for each x in the regression.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "SLOPE",
    description: "Returns the slope of the linear regression line through the given data points.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "INTERCEPT",
    description: "Returns the intercept of the linear regression line through the given data points.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "FORECAST.LINEAR",
    description: "Returns a value along a linear trend.",
    args: [
      { name: "x", type: "number" },
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "FORECAST",
    description: "Returns a value along a linear trend (legacy).",
    args: [
      { name: "x", type: "number" },
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "FORECAST.ETS",
    description: "Returns a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing algorithm.",
    args: [
      { name: "target_date", type: "value" },
      { name: "values", type: "range" },
      { name: "timeline", type: "range" },
      { name: "seasonality", type: "number", optional: true },
      { name: "data_completion", type: "number", optional: true },
      { name: "aggregation", type: "number", optional: true },
    ],
  },
  {
    name: "FORECAST.ETS.CONFINT",
    description: "Returns a confidence interval for a forecast value.",
    args: [
      { name: "target_date", type: "value" },
      { name: "values", type: "range" },
      { name: "timeline", type: "range" },
      { name: "confidence_level", type: "number" },
      { name: "seasonality", type: "number", optional: true },
      { name: "data_completion", type: "number", optional: true },
      { name: "aggregation", type: "number", optional: true },
    ],
  },
  {
    name: "FORECAST.ETS.SEASONALITY",
    description: "Returns the length of the repetitive pattern Excel detects for the specified time series.",
    args: [
      { name: "values", type: "range" },
      { name: "timeline", type: "range" },
      { name: "data_completion", type: "number", optional: true },
      { name: "aggregation", type: "number", optional: true },
    ],
  },
  {
    name: "FORECAST.ETS.STAT",
    description: "Returns a statistical value as the result of time series forecasting.",
    args: [
      { name: "values", type: "range" },
      { name: "timeline", type: "range" },
      { name: "statistic_type", type: "number" },
      { name: "seasonality", type: "number", optional: true },
      { name: "data_completion", type: "number", optional: true },
      { name: "aggregation", type: "number", optional: true },
    ],
  },
  {
    name: "LOGEST",
    description: "Returns the statistics for an exponential curve fit.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range", optional: true },
      { name: "const", type: "boolean", optional: true },
      { name: "stats", type: "boolean", optional: true },
    ],
  },
  {
    name: "LINEST",
    description: "Returns the statistics for a linear trend by using the least squares method.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range", optional: true },
      { name: "const", type: "boolean", optional: true },
      { name: "stats", type: "boolean", optional: true },
    ],
  },
  {
    name: "TREND",
    description: "Returns values along a linear trend.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range", optional: true },
      { name: "new_xs", type: "range", optional: true },
      { name: "const", type: "boolean", optional: true },
    ],
  },
  {
    name: "GROWTH",
    description: "Returns values along an exponential trend.",
    args: [
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range", optional: true },
      { name: "new_xs", type: "range", optional: true },
      { name: "const", type: "boolean", optional: true },
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
    name: "SUMSQ",
    description: "Returns the sum of the squares of the arguments.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "SUMX2MY2",
    description: "Returns the sum of the difference of squares of corresponding values in two arrays.",
    args: [
      { name: "array_x", type: "range" },
      { name: "array_y", type: "range" },
    ],
  },
  {
    name: "SUMX2PY2",
    description: "Returns the sum of the sum of squares of corresponding values in two arrays.",
    args: [
      { name: "array_x", type: "range" },
      { name: "array_y", type: "range" },
    ],
  },
  {
    name: "SUMXMY2",
    description: "Returns the sum of squares of differences of corresponding values in two arrays.",
    args: [
      { name: "array_x", type: "range" },
      { name: "array_y", type: "range" },
    ],
  },
  {
    name: "PRODUCT",
    description: "Multiplies all the numbers given as arguments and returns the product.",
    args: [
      { name: "number1", type: "range", repeating: true },
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
    name: "OFFSET",
    description: "Returns a reference offset from a given reference.",
    args: [
      { name: "reference", type: "range" },
      { name: "rows", type: "number" },
      { name: "cols", type: "number" },
      { name: "height", type: "number", optional: true },
      { name: "width", type: "number", optional: true },
    ],
  },
  {
    name: "ROW",
    description: "Returns the row number of a reference.",
    args: [
      { name: "reference", type: "range", optional: true },
    ],
  },
  {
    name: "COLUMN",
    description: "Returns the column number of a reference.",
    args: [
      { name: "reference", type: "range", optional: true },
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
    name: "FVSCHEDULE",
    description: "Returns the future value of an initial principal after applying a series of compound interest rates.",
    args: [
      { name: "principal", type: "number" },
      { name: "schedule", type: "range" },
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
    name: "MIRR",
    description: "Returns the modified internal rate of return for a series of periodic cash flows.",
    args: [
      { name: "values", type: "range" },
      { name: "finance_rate", type: "number" },
      { name: "reinvest_rate", type: "number" },
    ],
  },
  {
    name: "ACCRINT",
    description: "Returns the accrued interest for a security that pays periodic interest.",
    args: [
      { name: "issue", type: "value" },
      { name: "first_interest", type: "value" },
      { name: "settlement", type: "value" },
      { name: "rate", type: "number" },
      { name: "par", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
      { name: "calc_method", type: "boolean", optional: true },
    ],
  },
  {
    name: "ACCRINTM",
    description: "Returns the accrued interest for a security that pays interest at maturity.",
    args: [
      { name: "issue", type: "value" },
      { name: "settlement", type: "value" },
      { name: "rate", type: "number" },
      { name: "par", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "PRICE",
    description: "Returns the price per $100 face value of a security that pays periodic interest.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "rate", type: "number" },
      { name: "yld", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "YIELD",
    description: "Returns the yield on a security that pays periodic interest.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "rate", type: "number" },
      { name: "pr", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "DURATION",
    description: "Returns the annual duration of a security with periodic interest payments.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "coupon", type: "number" },
      { name: "yld", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "MDURATION",
    description: "Returns the modified Macauley duration for a security with periodic interest payments.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "coupon", type: "number" },
      { name: "yld", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "PRICEDISC",
    description: "Returns the price per $100 face value of a discounted security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "discount", type: "number" },
      { name: "redemption", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "PRICEMAT",
    description: "Returns the price per $100 face value of a security that pays interest at maturity.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "issue", type: "value" },
      { name: "rate", type: "number" },
      { name: "yld", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "YIELDDISC",
    description: "Returns the annual yield for a discounted security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "number" },
      { name: "redemption", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "YIELDMAT",
    description: "Returns the annual yield for a security that pays interest at maturity.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "issue", type: "value" },
      { name: "rate", type: "number" },
      { name: "pr", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "DISC",
    description: "Returns the discount rate for a security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "number" },
      { name: "redemption", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "RECEIVED",
    description: "Returns the amount received at maturity for a fully invested security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "investment", type: "number" },
      { name: "discount", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "INTRATE",
    description: "Returns the interest rate for a fully invested security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "investment", type: "number" },
      { name: "redemption", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPDAYBS",
    description: "Returns the number of days from the beginning of the coupon period to the settlement date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPDAYS",
    description: "Returns the number of days in the coupon period that contains the settlement date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPDAYSNC",
    description: "Returns the number of days from the settlement date to the next coupon date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPNCD",
    description: "Returns the next coupon date after the settlement date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPNUM",
    description: "Returns the number of coupons payable between the settlement date and maturity date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "COUPPCD",
    description: "Returns the previous coupon date before the settlement date.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "TBILLEQ",
    description: "Returns the bond-equivalent yield for a Treasury bill.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "discount", type: "number" },
    ],
  },
  {
    name: "TBILLPRICE",
    description: "Returns the price per $100 face value for a Treasury bill.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "discount", type: "number" },
    ],
  },
  {
    name: "TBILLYIELD",
    description: "Returns the yield for a Treasury bill.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "number" },
    ],
  },
  {
    name: "ODDLPRICE",
    description: "Returns the price per $100 face value of a security with an odd last period.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "last_interest", type: "value" },
      { name: "rate", type: "number" },
      { name: "yld", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "ODDLYIELD",
    description: "Returns the yield of a security that has an odd last period.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "last_interest", type: "value" },
      { name: "rate", type: "number" },
      { name: "pr", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "ODDFPRICE",
    description: "Returns the price per $100 face value of a security with an odd first period.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "issue", type: "value" },
      { name: "first_interest", type: "value" },
      { name: "rate", type: "number" },
      { name: "yld", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "ODDFYIELD",
    description: "Returns the yield of a security with an odd first period.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "issue", type: "value" },
      { name: "first_interest", type: "value" },
      { name: "rate", type: "number" },
      { name: "pr", type: "number" },
      { name: "redemption", type: "number" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "AMORDEGRC",
    description: "Returns the depreciation for each accounting period by using a depreciation coefficient.",
    args: [
      { name: "cost", type: "number" },
      { name: "date_purchased", type: "value" },
      { name: "first_period", type: "value" },
      { name: "salvage", type: "number" },
      { name: "period", type: "number" },
      { name: "rate", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "AMORLINC",
    description: "Returns the depreciation for each accounting period.",
    args: [
      { name: "cost", type: "number" },
      { name: "date_purchased", type: "value" },
      { name: "first_period", type: "value" },
      { name: "salvage", type: "number" },
      { name: "period", type: "number" },
      { name: "rate", type: "number" },
      { name: "basis", type: "number", optional: true },
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
    name: "AND",
    description: "Returns TRUE if all arguments evaluate to TRUE.",
    args: [
      { name: "logical1", type: "value", repeating: true },
    ],
  },
  {
    name: "OR",
    description: "Returns TRUE if any argument evaluates to TRUE.",
    args: [
      { name: "logical1", type: "value", repeating: true },
    ],
  },
  {
    name: "XOR",
    description: "Returns a logical exclusive OR of all arguments.",
    args: [
      { name: "logical1", type: "value", repeating: true },
    ],
  },
  {
    name: "NOT",
    description: "Reverses the value of its argument.",
    args: [
      { name: "logical", type: "value" },
    ],
  },
  {
    name: "CHOOSE",
    description: "Chooses a value from a list based on an index number.",
    args: [
      { name: "index_num", type: "number" },
      { name: "value1", type: "value", repeating: true },
    ],
  },
  {
    name: "ISBLANK",
    description: "Returns TRUE if the value is blank.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISERR",
    description: "Returns TRUE if the value is any error value except #N/A.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISERROR",
    description: "Returns TRUE if the value is any error value.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISLOGICAL",
    description: "Returns TRUE if the value is a logical value.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISNA",
    description: "Returns TRUE if the value is the #N/A error value.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISNONTEXT",
    description: "Returns TRUE if the value is not text.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISNUMBER",
    description: "Returns TRUE if the value is a number.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISOMITTED",
    description: "Returns TRUE if the argument is omitted in a LAMBDA call.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISREF",
    description: "Returns TRUE if the value is a reference.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ISTEXT",
    description: "Returns TRUE if the value is text.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "N",
    description: "Converts a value to a number.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "T",
    description: "Converts its arguments to text.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "TYPE",
    description: "Returns a number indicating the data type of a value.",
    args: [
      { name: "value", type: "value" },
    ],
  },
  {
    name: "ERROR.TYPE",
    description: "Returns a number corresponding to an error type.",
    args: [
      { name: "error_val", type: "value" },
    ],
  },
  {
    name: "IFS",
    description: "Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.",
    args: [
      { name: "logical_test1", type: "value", repeating: true },
      { name: "value_if_true1", type: "value" },
    ],
  },
  {
    name: "SWITCH",
    description: "Evaluates an expression against a list of values and returns the result corresponding to the first match.",
    args: [
      { name: "expression", type: "value" },
      { name: "value1", type: "value", repeating: true },
      { name: "result1", type: "value" },
    ],
  },
  {
    name: "LET",
    description: "Assigns names to calculation results to improve readability and performance.",
    args: [
      { name: "name1", type: "string" },
      { name: "value1", type: "value" },
      { name: "calculation", type: "value" },
    ],
  },
  {
    name: "LAMBDA",
    description: "Creates a custom, reusable function and returns it as a value.",
    args: [
      { name: "parameter1", type: "string", optional: true },
      { name: "calculation", type: "value" },
    ],
  },
  {
    name: "IFERROR",
    description: "Returns a value you specify if a formula evaluates to an error; otherwise returns the formula result.",
    args: [
      { name: "value", type: "value" },
      { name: "value_if_error", type: "value" },
    ],
  },
  {
    name: "IFNA",
    description: "Returns a value you specify if a formula evaluates to the #N/A error; otherwise returns the formula result.",
    args: [
      { name: "value", type: "value" },
      { name: "value_if_na", type: "value" },
    ],
  },
  {
    name: "LEFT",
    description: "Returns the leftmost characters from a text value.",
    args: [
      { name: "text", type: "range" },
      { name: "num_chars", type: "number", optional: true },
    ],
  },
  {
    name: "RIGHT",
    description: "Returns the rightmost characters from a text value.",
    args: [
      { name: "text", type: "range" },
      { name: "num_chars", type: "number", optional: true },
    ],
  },
  {
    name: "MID",
    description: "Returns a specific number of characters from a text string, starting at the position you specify.",
    args: [
      { name: "text", type: "range" },
      { name: "start_num", type: "number" },
      { name: "num_chars", type: "number" },
    ],
  },
  {
    name: "LEN",
    description: "Returns the number of characters in a text string.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "TRIM",
    description: "Removes all spaces from text except for single spaces between words.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "CLEAN",
    description: "Removes all nonprintable characters from text.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "LOWER",
    description: "Converts text to lowercase.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "UPPER",
    description: "Converts text to uppercase.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "PROPER",
    description: "Capitalizes the first letter in each word of a text value.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "SUBSTITUTE",
    description: "Substitutes new text for old text in a text string.",
    args: [
      { name: "text", type: "range" },
      { name: "old_text", type: "value" },
      { name: "new_text", type: "value" },
      { name: "instance_num", type: "number", optional: true },
    ],
  },
  {
    name: "REPLACE",
    description: "Replaces part of a text string with a different text string.",
    args: [
      { name: "old_text", type: "range" },
      { name: "start_num", type: "number" },
      { name: "num_chars", type: "number" },
      { name: "new_text", type: "value" },
    ],
  },
  {
    name: "FIND",
    description: "Finds one text value within another (case-sensitive).",
    args: [
      { name: "find_text", type: "value" },
      { name: "within_text", type: "range" },
      { name: "start_num", type: "number", optional: true },
    ],
  },
  {
    name: "SEARCH",
    description: "Finds one text value within another (case-insensitive) and returns the position.",
    args: [
      { name: "find_text", type: "value" },
      { name: "within_text", type: "range" },
      { name: "start_num", type: "number", optional: true },
    ],
  },
  {
    name: "VALUE",
    description: "Converts a text string that represents a number to a number.",
    args: [
      { name: "text", type: "range" },
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
    name: "ROWS",
    description: "Returns the number of rows in a reference or array.",
    args: [
      { name: "array", type: "range" },
    ],
  },
  {
    name: "COLUMNS",
    description: "Returns the number of columns in a reference or array.",
    args: [
      { name: "array", type: "range" },
    ],
  },
  {
    name: "MMULT",
    description: "Returns the matrix product of two arrays.",
    args: [
      { name: "array1", type: "range" },
      { name: "array2", type: "range" },
    ],
  },
  {
    name: "MDETERM",
    description: "Returns the matrix determinant of an array.",
    args: [
      { name: "array", type: "range" },
    ],
  },
  {
    name: "MINVERSE",
    description: "Returns the inverse matrix for a matrix stored in an array.",
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
    name: "HSTACK",
    description: "Appends arrays horizontally and in sequence to return a larger array.",
    args: [
      { name: "array1", type: "range", repeating: true },
    ],
  },
  {
    name: "VSTACK",
    description: "Appends arrays vertically and in sequence to return a larger array.",
    args: [
      { name: "array1", type: "range", repeating: true },
    ],
  },
  {
    name: "TOCOL",
    description: "Returns a single column containing all the items in the specified array or range.",
    args: [
      { name: "array", type: "range" },
      { name: "ignore", type: "number", optional: true },
      { name: "scan_by_column", type: "boolean", optional: true },
    ],
  },
  {
    name: "TOROW",
    description: "Returns a single row containing all the items in the specified array or range.",
    args: [
      { name: "array", type: "range" },
      { name: "ignore", type: "number", optional: true },
      { name: "scan_by_column", type: "boolean", optional: true },
    ],
  },
  {
    name: "WRAPROWS",
    description: "Wraps a row or column of values by rows after a specified number of elements.",
    args: [
      { name: "vector", type: "range" },
      { name: "wrap_count", type: "number" },
      { name: "pad_with", type: "value", optional: true },
    ],
  },
  {
    name: "WRAPCOLS",
    description: "Wraps a row or column of values by columns after a specified number of elements.",
    args: [
      { name: "vector", type: "range" },
      { name: "wrap_count", type: "number" },
      { name: "pad_with", type: "value", optional: true },
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
  {
    name: "MAP",
    description: "Maps a LAMBDA to one or more arrays and returns an array of results.",
    args: [
      { name: "array", type: "range" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "REDUCE",
    description: "Reduces an array to an accumulated value by applying a LAMBDA to each value.",
    args: [
      { name: "initial_value", type: "range", optional: true },
      { name: "array", type: "range" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "SCAN",
    description: "Scans an array by applying a LAMBDA to each value and returns intermediate results as an array.",
    args: [
      { name: "initial_value", type: "range", optional: true },
      { name: "array", type: "range" },
      { name: "lambda", type: "value" },
    ],
  },
  {
    name: "AREAS",
    description: "Returns the number of areas in a reference.",
    args: [
      { name: "reference", type: "range" },
    ],
  },
  {
    name: "CELL",
    description: "Returns information about the formatting, location, or contents of a cell.",
    args: [
      { name: "info_type", type: "string" },
      { name: "reference", type: "range", optional: true },
    ],
  },
  {
    name: "FORMULATEXT",
    description: "Returns the formula in a given cell as text.",
    args: [
      { name: "reference", type: "range" },
    ],
  },
  {
    name: "ISFORMULA",
    description: "Returns TRUE if there is a formula in a referenced cell.",
    args: [
      { name: "reference", type: "range" },
    ],
  },
  {
    name: "GETPIVOTDATA",
    description: "Returns data stored in a PivotTable report.",
    args: [
      { name: "data_field", type: "string" },
      { name: "pivot_table", type: "range" },
      { name: "field1", type: "string", repeating: true },
      { name: "item1", type: "string" },
    ],
  },
  {
    name: "ADDRESS",
    description: "Creates a cell reference as text, given specified row and column numbers.",
    args: [
      { name: "row_num", type: "number" },
      { name: "column_num", type: "number" },
      { name: "abs_num", type: "number", optional: true },
      { name: "a1", type: "boolean", optional: true },
      { name: "sheet_text", type: "string", optional: true },
    ],
  },
  {
    name: "INDIRECT",
    description: "Returns the reference specified by a text string.",
    args: [
      { name: "ref_text", type: "string" },
      { name: "a1", type: "boolean", optional: true },
    ],
  },
  {
    name: "INFO",
    description: "Returns information about the current operating environment.",
    args: [
      { name: "type_text", type: "string" },
    ],
  },
  {
    name: "IMAGE",
    description: "Inserts an image from a given URL or data source.",
    args: [
      { name: "source", type: "string" },
      { name: "alt_text", type: "string", optional: true },
      { name: "sizing", type: "number", optional: true },
      { name: "height", type: "number", optional: true },
      { name: "width", type: "number", optional: true },
    ],
  },
  {
    name: "NUMBERVALUE",
    description: "Converts text to a number in a locale-independent way, using custom separators.",
    args: [
      { name: "text", type: "value" },
      { name: "decimal_separator", type: "string", optional: true },
      { name: "group_separator", type: "string", optional: true },
    ],
  },
  {
    name: "SHEET",
    description: "Returns the sheet number of the referenced sheet.",
    args: [
      { name: "value", type: "range", optional: true },
    ],
  },
  {
    name: "HYPERLINK",
    description: "Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet.",
    args: [
      { name: "link_location", type: "string" },
      { name: "friendly_name", type: "string", optional: true },
    ],
  },
  {
    name: "SHEETS",
    description: "Returns the number of sheets in a reference.",
    args: [
      { name: "reference", type: "range", optional: true },
    ],
  },
  {
    name: "TEXTAFTER",
    description: "Returns text that occurs after a given delimiter.",
    minArgs: 2,
    maxArgs: 6,
    args: [
      { name: "text", type: "range" },
      { name: "delimiter", type: "string" },
      { name: "instance_num", type: "number", optional: true },
      { name: "match_mode", type: "number", optional: true },
      { name: "match_end", type: "number", optional: true },
      { name: "if_not_found", type: "value", optional: true },
    ],
  },
  {
    name: "TEXTBEFORE",
    description: "Returns text that occurs before a given delimiter.",
    minArgs: 2,
    maxArgs: 6,
    args: [
      { name: "text", type: "range" },
      { name: "delimiter", type: "string" },
      { name: "instance_num", type: "number", optional: true },
      { name: "match_mode", type: "number", optional: true },
      { name: "match_end", type: "number", optional: true },
      { name: "if_not_found", type: "value", optional: true },
    ],
  },
  {
    name: "TEXT",
    description: "Formats a number and converts it to text.",
    args: [
      { name: "value", type: "value" },
      { name: "format_text", type: "string" },
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
