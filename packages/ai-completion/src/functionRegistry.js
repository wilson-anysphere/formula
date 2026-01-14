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
    name: "ABS",
    description: "Returns the absolute value of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ROUND",
    description: "Rounds a number to a specified number of digits.",
    args: [
      { name: "number", type: "value" },
      { name: "num_digits", type: "number" },
    ],
  },
  {
    name: "ROUNDUP",
    description: "Rounds a number up, away from zero.",
    args: [
      { name: "number", type: "value" },
      { name: "num_digits", type: "number" },
    ],
  },
  {
    name: "ROUNDDOWN",
    description: "Rounds a number down, toward zero.",
    args: [
      { name: "number", type: "value" },
      { name: "num_digits", type: "number" },
    ],
  },
  {
    name: "INT",
    description: "Rounds a number down to the nearest integer.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "TRUNC",
    description: "Truncates a number to an integer by removing the fractional part.",
    args: [
      { name: "number", type: "value" },
      { name: "num_digits", type: "number", optional: true },
    ],
  },
  {
    name: "POWER",
    description: "Returns the result of a number raised to a power.",
    args: [
      { name: "number", type: "value" },
      { name: "power", type: "number" },
    ],
  },
  {
    name: "SQRT",
    description: "Returns the positive square root of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "EXP",
    description: "Returns e raised to the power of a given number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "LN",
    description: "Returns the natural logarithm of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "LOG",
    description: "Returns the logarithm of a number to a specified base.",
    args: [
      { name: "number", type: "value" },
      { name: "base", type: "number", optional: true },
    ],
  },
  {
    name: "LOG10",
    description: "Returns the base-10 logarithm of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "MOD",
    description: "Returns the remainder after number is divided by divisor.",
    args: [
      { name: "number", type: "value" },
      { name: "divisor", type: "number" },
    ],
  },
  {
    name: "BITAND",
    description: "Returns a bitwise AND of its arguments.",
    args: [
      { name: "number1", type: "value" },
      { name: "number2", type: "value" },
    ],
  },
  {
    name: "BITOR",
    description: "Returns a bitwise OR of its arguments.",
    args: [
      { name: "number1", type: "value" },
      { name: "number2", type: "value" },
    ],
  },
  {
    name: "BITXOR",
    description: "Returns a bitwise XOR of its arguments.",
    args: [
      { name: "number1", type: "value" },
      { name: "number2", type: "value" },
    ],
  },
  {
    name: "BITLSHIFT",
    description: "Returns a number shifted left by a specified number of bits.",
    args: [
      { name: "number", type: "value" },
      { name: "shift_amount", type: "number" },
    ],
  },
  {
    name: "BITRSHIFT",
    description: "Returns a number shifted right by a specified number of bits.",
    args: [
      { name: "number", type: "value" },
      { name: "shift_amount", type: "number" },
    ],
  },
  {
    name: "QUOTIENT",
    description: "Returns the integer portion of a division.",
    args: [
      { name: "numerator", type: "value" },
      { name: "denominator", type: "value" },
    ],
  },
  {
    name: "FACT",
    description: "Returns the factorial of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "FACTDOUBLE",
    description: "Returns the double factorial of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "COMBIN",
    description: "Returns the number of combinations for a given number of items.",
    args: [
      { name: "number", type: "value" },
      { name: "number_chosen", type: "value" },
    ],
  },
  {
    name: "COMBINA",
    description: "Returns the number of combinations with repetitions for a given number of items.",
    args: [
      { name: "number", type: "value" },
      { name: "number_chosen", type: "value" },
    ],
  },
  {
    name: "PERMUT",
    description: "Returns the number of permutations for a given number of objects.",
    args: [
      { name: "number", type: "value" },
      { name: "number_chosen", type: "value" },
    ],
  },
  {
    name: "PERMUTATIONA",
    description: "Returns the number of permutations with repetitions for a given number of objects.",
    args: [
      { name: "number", type: "value" },
      { name: "number_chosen", type: "value" },
    ],
  },
  {
    name: "GAMMA",
    description: "Returns the Gamma function value.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "GAMMALN",
    description: "Returns the natural logarithm of the Gamma function.",
    args: [
      { name: "x", type: "value" },
    ],
  },
  {
    name: "GAMMALN.PRECISE",
    description: "Returns the natural logarithm of the Gamma function, for positive numbers.",
    args: [
      { name: "x", type: "value" },
    ],
  },
  {
    name: "GAUSS",
    description: "Returns the probability that a standard normal random variable is less than z.",
    args: [
      { name: "z", type: "value" },
    ],
  },
  {
    name: "ERF",
    description: "Returns the error function integrated between lower_limit and upper_limit.",
    args: [
      { name: "lower_limit", type: "value" },
      { name: "upper_limit", type: "value", optional: true },
    ],
  },
  {
    name: "ERFC",
    description: "Returns the complementary error function.",
    args: [
      { name: "x", type: "value" },
    ],
  },
  {
    name: "BESSELI",
    description: "Returns the modified Bessel function In(x).",
    args: [
      { name: "x", type: "value" },
      { name: "n", type: "number" },
    ],
  },
  {
    name: "BESSELJ",
    description: "Returns the Bessel function Jn(x).",
    args: [
      { name: "x", type: "value" },
      { name: "n", type: "number" },
    ],
  },
  {
    name: "BESSELK",
    description: "Returns the modified Bessel function Kn(x).",
    args: [
      { name: "x", type: "value" },
      { name: "n", type: "number" },
    ],
  },
  {
    name: "BESSELY",
    description: "Returns the Bessel function Yn(x).",
    args: [
      { name: "x", type: "value" },
      { name: "n", type: "number" },
    ],
  },
  {
    name: "DELTA",
    description: "Tests whether two values are equal.",
    args: [
      { name: "number1", type: "value" },
      { name: "number2", type: "number", optional: true },
    ],
  },
  {
    name: "GESTEP",
    description: "Tests whether a number is greater than or equal to a threshold.",
    args: [
      { name: "number", type: "value" },
      { name: "step", type: "number", optional: true },
    ],
  },
  {
    name: "SIGN",
    description: "Returns the sign of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "DEGREES",
    description: "Converts radians to degrees.",
    args: [
      { name: "angle", type: "value" },
    ],
  },
  {
    name: "RADIANS",
    description: "Converts degrees to radians.",
    args: [
      { name: "angle", type: "value" },
    ],
  },
  {
    name: "SIN",
    description: "Returns the sine of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "COS",
    description: "Returns the cosine of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "TAN",
    description: "Returns the tangent of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "SINH",
    description: "Returns the hyperbolic sine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "COSH",
    description: "Returns the hyperbolic cosine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "TANH",
    description: "Returns the hyperbolic tangent of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "CSC",
    description: "Returns the cosecant of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "SEC",
    description: "Returns the secant of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "COT",
    description: "Returns the cotangent of an angle.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "CSCH",
    description: "Returns the hyperbolic cosecant of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "SECH",
    description: "Returns the hyperbolic secant of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "COTH",
    description: "Returns the hyperbolic cotangent of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ASIN",
    description: "Returns the arcsine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ASINH",
    description: "Returns the inverse hyperbolic sine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ACOS",
    description: "Returns the arccosine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ACOSH",
    description: "Returns the inverse hyperbolic cosine of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ATAN",
    description: "Returns the arctangent of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ATANH",
    description: "Returns the inverse hyperbolic tangent of a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ATAN2",
    description: "Returns the arctangent from x- and y-coordinates.",
    args: [
      { name: "x_num", type: "value" },
      { name: "y_num", type: "value" },
    ],
  },
  {
    name: "ACOT",
    description: "Returns the principal value of the arccotangent of a number.",
    args: [
      { name: "number", type: "value" },
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
    name: "GCD",
    description: "Returns the greatest common divisor.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "LCM",
    description: "Returns the least common multiple.",
    args: [
      { name: "number1", type: "range", repeating: true },
    ],
  },
  {
    name: "MULTINOMIAL",
    description: "Returns the multinomial of a set of numbers.",
    args: [
      { name: "number1", type: "range", repeating: true },
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
    name: "CEILING.MATH",
    description: "Rounds a number up to the nearest integer or nearest multiple of significance.",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number", optional: true },
      { name: "mode", type: "number", optional: true },
    ],
  },
  {
    name: "FLOOR.MATH",
    description: "Rounds a number down to the nearest integer or nearest multiple of significance.",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number", optional: true },
      { name: "mode", type: "number", optional: true },
    ],
  },
  {
    name: "CEILING",
    description: "Rounds a number up, away from zero, to the nearest multiple of significance.",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number" },
    ],
  },
  {
    name: "FLOOR",
    description: "Rounds a number down, toward zero, to the nearest multiple of significance.",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number" },
    ],
  },
  {
    name: "CEILING.PRECISE",
    description: "Rounds a number up to the nearest integer or multiple of significance (always toward +∞).",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "FLOOR.PRECISE",
    description: "Rounds a number down to the nearest integer or multiple of significance (always toward -∞).",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "ISO.CEILING",
    description: "Rounds a number up to the nearest integer or multiple of significance (ISO variant).",
    args: [
      { name: "number", type: "range" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "MROUND",
    description: "Rounds a number to the nearest multiple of a specified value.",
    args: [
      { name: "number", type: "range" },
      { name: "multiple", type: "number" },
    ],
  },
  {
    name: "EVEN",
    description: "Rounds a number up to the nearest even integer.",
    args: [
      { name: "number", type: "range" },
    ],
  },
  {
    name: "ODD",
    description: "Rounds a number up to the nearest odd integer.",
    args: [
      { name: "number", type: "range" },
    ],
  },
  {
    name: "WORKDAY",
    description: "Returns a date that is a specified number of working days before or after a start date.",
    args: [
      { name: "start_date", type: "value" },
      { name: "days", type: "value" },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "DATE",
    description: "Returns the serial number of a particular date.",
    args: [
      { name: "year", type: "value" },
      { name: "month", type: "value" },
      { name: "day", type: "value" },
    ],
  },
  {
    name: "TIME",
    description: "Returns the serial number of a particular time.",
    args: [
      { name: "hour", type: "value" },
      { name: "minute", type: "value" },
      { name: "second", type: "value" },
    ],
  },
  {
    name: "YEAR",
    description: "Returns the year corresponding to a date.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "MONTH",
    description: "Returns the month corresponding to a date.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "DAY",
    description: "Returns the day of the month corresponding to a date.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "EDATE",
    description: "Returns the serial number of the date that is the indicated number of months before or after a specified date.",
    args: [
      { name: "start_date", type: "value" },
      { name: "months", type: "value" },
    ],
  },
  {
    name: "EOMONTH",
    description: "Returns the serial number for the last day of the month that is the indicated number of months before or after a specified date.",
    args: [
      { name: "start_date", type: "value" },
      { name: "months", type: "value" },
    ],
  },
  {
    name: "WEEKDAY",
    description: "Converts a serial number to a day of the week.",
    args: [
      { name: "serial_number", type: "value" },
      { name: "return_type", type: "number", optional: true },
    ],
  },
  {
    name: "WEEKNUM",
    description: "Converts a serial number to a number representing where the week falls numerically with a year.",
    args: [
      { name: "serial_number", type: "value" },
      { name: "return_type", type: "number", optional: true },
    ],
  },
  {
    name: "ISO.WEEKNUM",
    description: "Returns the ISO week number for the given date.",
    args: [
      { name: "serial_number", type: "value" },
    ],
  },
  {
    name: "ISOWEEKNUM",
    description: "Returns the ISO week number for the given date (legacy name).",
    args: [
      { name: "serial_number", type: "value" },
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
      { name: "date_text", type: "range" },
    ],
  },
  {
    name: "TIMEVALUE",
    description: "Converts a time in the form of text to a serial number.",
    args: [
      { name: "time_text", type: "range" },
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
      { name: "start_date", type: "value" },
      { name: "days", type: "value" },
      { name: "weekend", type: "number", optional: true },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "NETWORKDAYS",
    description: "Returns the number of whole working days between two dates.",
    args: [
      { name: "start_date", type: "value" },
      { name: "end_date", type: "value" },
      { name: "holidays", type: "range", optional: true },
    ],
  },
  {
    name: "NETWORKDAYS.INTL",
    description: "Returns the number of whole working days between two dates with custom weekends.",
    args: [
      { name: "start_date", type: "value" },
      { name: "end_date", type: "value" },
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
      { name: "number", type: "value" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "RANK",
    description: "Returns the rank of a number in a list of numbers (legacy).",
    args: [
      { name: "number", type: "value" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "RANK.AVG",
    description: "Returns the rank of a number in a list of numbers, with ties averaged.",
    args: [
      { name: "number", type: "value" },
      { name: "ref", type: "range" },
      { name: "order", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK",
    description: "Returns the rank of a value in a data set as a percentage of the data set.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "value" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK.INC",
    description: "Returns the rank of a value in a data set as a percentage of the data set, inclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "value" },
      { name: "significance", type: "number", optional: true },
    ],
  },
  {
    name: "PERCENTRANK.EXC",
    description: "Returns the rank of a value in a data set as a percentage of the data set, exclusive.",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "value" },
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
      { name: "percent", type: "value" },
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
      { name: "x", type: "value" },
      { name: "sigma", type: "value", optional: true },
    ],
  },
  {
    name: "ZTEST",
    description: "Returns the one-tailed probability-value of a z-test (legacy).",
    args: [
      { name: "array", type: "range" },
      { name: "x", type: "value" },
      { name: "sigma", type: "value", optional: true },
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
    name: "FISHER",
    description: "Returns the Fisher transformation.",
    args: [
      { name: "x", type: "value" },
    ],
  },
  {
    name: "FISHERINV",
    description: "Returns the inverse of the Fisher transformation.",
    args: [
      { name: "y", type: "value" },
    ],
  },
  {
    name: "PHI",
    description: "Returns the value of the standard normal density function.",
    args: [
      { name: "x", type: "value" },
    ],
  },
  {
    name: "NORM.DIST",
    description: "Returns the normal cumulative distribution for the specified mean and standard deviation.",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "NORM.S.DIST",
    description: "Returns the standard normal cumulative distribution function.",
    args: [
      { name: "z", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "NORM.INV",
    description: "Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.",
    args: [
      { name: "probability", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
    ],
  },
  {
    name: "NORM.S.INV",
    description: "Returns the inverse of the standard normal cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
    ],
  },
  {
    name: "BINOM.DIST",
    description: "Returns the individual term binomial distribution probability.",
    args: [
      { name: "number_s", type: "value" },
      { name: "trials", type: "value" },
      { name: "probability_s", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "BINOMDIST",
    description: "Returns the individual term binomial distribution probability (legacy).",
    args: [
      { name: "number_s", type: "value" },
      { name: "trials", type: "value" },
      { name: "probability_s", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "BINOM.DIST.RANGE",
    description: "Returns the probability of a trial result using a binomial distribution.",
    args: [
      { name: "trials", type: "value" },
      { name: "probability_s", type: "value" },
      { name: "number_s", type: "value" },
      { name: "number_s2", type: "value", optional: true },
    ],
  },
  {
    name: "BINOM.INV",
    description: "Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value.",
    args: [
      { name: "trials", type: "value" },
      { name: "probability_s", type: "value" },
      { name: "alpha", type: "value" },
    ],
  },
  {
    name: "HYPGEOM.DIST",
    description: "Returns the hypergeometric distribution.",
    args: [
      { name: "sample_s", type: "value" },
      { name: "number_sample", type: "value" },
      { name: "population_s", type: "value" },
      { name: "number_pop", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "HYPGEOMDIST",
    description: "Returns the hypergeometric distribution (legacy).",
    args: [
      { name: "sample_s", type: "value" },
      { name: "number_sample", type: "value" },
      { name: "population_s", type: "value" },
      { name: "number_pop", type: "value" },
    ],
  },
  {
    name: "NEGBINOM.DIST",
    description: "Returns the negative binomial distribution.",
    args: [
      { name: "number_f", type: "value" },
      { name: "number_s", type: "value" },
      { name: "probability_s", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "NEGBINOMDIST",
    description: "Returns the negative binomial distribution (legacy).",
    args: [
      { name: "number_f", type: "value" },
      { name: "number_s", type: "value" },
      { name: "probability_s", type: "value" },
    ],
  },
  {
    name: "NORMDIST",
    description: "Returns the normal cumulative distribution for the specified mean and standard deviation (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "NORMSDIST",
    description: "Returns the standard normal cumulative distribution function (legacy).",
    args: [
      { name: "z", type: "value" },
    ],
  },
  {
    name: "NORMINV",
    description: "Returns the inverse of the normal cumulative distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
    ],
  },
  {
    name: "NORMSINV",
    description: "Returns the inverse of the standard normal cumulative distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
    ],
  },
  {
    name: "LOGNORM.DIST",
    description: "Returns the lognormal distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "LOGNORM.INV",
    description: "Returns the inverse of the lognormal cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
    ],
  },
  {
    name: "LOGNORMDIST",
    description: "Returns the lognormal cumulative distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "standard_dev", type: "value" },
    ],
  },
  {
    name: "EXPON.DIST",
    description: "Returns the exponential distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "lambda", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "EXPONDIST",
    description: "Returns the exponential distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "lambda", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "GAMMA.DIST",
    description: "Returns the gamma distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "GAMMADIST",
    description: "Returns the gamma distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "GAMMA.INV",
    description: "Returns the inverse of the gamma cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
    ],
  },
  {
    name: "GAMMAINV",
    description: "Returns the inverse of the gamma cumulative distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
    ],
  },
  {
    name: "BETA.DIST",
    description: "Returns the beta distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "cumulative", type: "boolean" },
      { name: "A", type: "value", optional: true },
      { name: "B", type: "value", optional: true },
    ],
  },
  {
    name: "BETA.INV",
    description: "Returns the inverse of the beta cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "A", type: "value", optional: true },
      { name: "B", type: "value", optional: true },
    ],
  },
  {
    name: "BETADIST",
    description: "Returns the beta cumulative distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "A", type: "value", optional: true },
      { name: "B", type: "value", optional: true },
    ],
  },
  {
    name: "BETAINV",
    description: "Returns the inverse of the beta cumulative distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "A", type: "value", optional: true },
      { name: "B", type: "value", optional: true },
    ],
  },
  {
    name: "CHISQ.DIST",
    description: "Returns the chi-squared distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "CHISQ.DIST.RT",
    description: "Returns the right-tailed probability of the chi-squared distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "CHISQ.INV",
    description: "Returns the inverse of the chi-squared cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "CHISQ.INV.RT",
    description: "Returns the inverse of the right-tailed probability of the chi-squared distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "CHIDIST",
    description: "Returns the right-tailed probability of the chi-squared distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "CHIINV",
    description: "Returns the inverse of the right-tailed probability of the chi-squared distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "F.DIST",
    description: "Returns the F distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "F.DIST.RT",
    description: "Returns the right-tailed probability of the F distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
    ],
  },
  {
    name: "F.INV",
    description: "Returns the inverse of the F cumulative distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
    ],
  },
  {
    name: "F.INV.RT",
    description: "Returns the inverse of the right-tailed probability of the F distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
    ],
  },
  {
    name: "FDIST",
    description: "Returns the (right-tailed) F distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
    ],
  },
  {
    name: "FINV",
    description: "Returns the inverse of the F distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom1", type: "value" },
      { name: "deg_freedom2", type: "value" },
    ],
  },
  {
    name: "T.DIST",
    description: "Returns the Student's t distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "T.DIST.2T",
    description: "Returns the two-tailed Student's t distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "T.DIST.RT",
    description: "Returns the right-tailed Student's t distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "T.INV",
    description: "Returns the inverse of the Student's t distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "T.INV.2T",
    description: "Returns the inverse of the two-tailed Student's t distribution.",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "TDIST",
    description: "Returns the Student's t distribution (legacy).",
    args: [
      { name: "x", type: "value" },
      { name: "deg_freedom", type: "value" },
      { name: "tails", type: "number" },
    ],
  },
  {
    name: "TINV",
    description: "Returns the inverse of the Student's t distribution (legacy).",
    args: [
      { name: "probability", type: "value" },
      { name: "deg_freedom", type: "value" },
    ],
  },
  {
    name: "POISSON",
    description: "Returns the Poisson distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "POISSON.DIST",
    description: "Returns the Poisson distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "mean", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "WEIBULL",
    description: "Returns the Weibull distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "WEIBULL.DIST",
    description: "Returns the Weibull distribution.",
    args: [
      { name: "x", type: "value" },
      { name: "alpha", type: "value" },
      { name: "beta", type: "value" },
      { name: "cumulative", type: "boolean" },
    ],
  },
  {
    name: "PROB",
    description: "Returns the probability that values in a range are between two limits.",
    args: [
      { name: "x_range", type: "range" },
      { name: "prob_range", type: "range" },
      { name: "lower_limit", type: "value" },
      { name: "upper_limit", type: "value", optional: true },
    ],
  },
  {
    name: "SERIESSUM",
    description: "Returns the sum of a power series based on the formula.",
    args: [
      { name: "x", type: "value" },
      { name: "n", type: "value" },
      { name: "m", type: "value" },
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
      { name: "x", type: "value" },
      { name: "known_ys", type: "range" },
      { name: "known_xs", type: "range" },
    ],
  },
  {
    name: "FORECAST",
    description: "Returns a value along a linear trend (legacy).",
    args: [
      { name: "x", type: "value" },
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
      // Excel accepts the optional args in the same order as the core ETS family:
      // seasonality, data_completion, aggregation, statistic_type.
      { name: "seasonality", type: "number", optional: true },
      { name: "data_completion", type: "number", optional: true },
      { name: "aggregation", type: "number", optional: true },
      { name: "statistic_type", type: "number", optional: true },
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
      { name: "rate", type: "value" },
      { name: "value1", type: "range", repeating: true },
    ],
  },
  {
    name: "PV",
    description: "Returns the present value of an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "nper", type: "value" },
      { name: "pmt", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "FV",
    description: "Returns the future value of an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "nper", type: "value" },
      { name: "pmt", type: "value" },
      { name: "pv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "PMT",
    description: "Returns the periodic payment for an annuity.",
    args: [
      { name: "rate", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "NPER",
    description: "Returns the number of periods for an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "pmt", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "RATE",
    description: "Returns the interest rate per period of an annuity.",
    args: [
      { name: "nper", type: "value" },
      { name: "pmt", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
      { name: "guess", type: "value", optional: true },
    ],
  },
  {
    name: "IPMT",
    description: "Returns the interest payment for a given period for an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "per", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "PPMT",
    description: "Returns the payment on the principal for a given period for an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "per", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value", optional: true },
      { name: "type", type: "number", optional: true },
    ],
  },
  {
    name: "CUMIPMT",
    description: "Returns the cumulative interest paid between two periods.",
    args: [
      { name: "rate", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "start_period", type: "value" },
      { name: "end_period", type: "value" },
      { name: "type", type: "number" },
    ],
  },
  {
    name: "CUMPRINC",
    description: "Returns the cumulative principal paid between two periods.",
    args: [
      { name: "rate", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "start_period", type: "value" },
      { name: "end_period", type: "value" },
      { name: "type", type: "number" },
    ],
  },
  {
    name: "ISPMT",
    description: "Calculates the interest paid during a specific period of an investment.",
    args: [
      { name: "rate", type: "value" },
      { name: "per", type: "value" },
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
    ],
  },
  {
    name: "RRI",
    description: "Returns an equivalent interest rate for the growth of an investment.",
    args: [
      { name: "nper", type: "value" },
      { name: "pv", type: "value" },
      { name: "fv", type: "value" },
    ],
  },
  {
    name: "EFFECT",
    description: "Returns the effective annual interest rate.",
    args: [
      { name: "nominal_rate", type: "value" },
      { name: "npery", type: "value" },
    ],
  },
  {
    name: "NOMINAL",
    description: "Returns the nominal annual interest rate.",
    args: [
      { name: "effect_rate", type: "value" },
      { name: "npery", type: "value" },
    ],
  },
  {
    name: "DOLLARDE",
    description: "Converts a dollar price expressed as a fraction into a decimal.",
    args: [
      { name: "fractional_dollar", type: "value" },
      { name: "fraction", type: "value" },
    ],
  },
  {
    name: "DOLLARFR",
    description: "Converts a dollar price expressed as a decimal into a fraction.",
    args: [
      { name: "decimal_dollar", type: "value" },
      { name: "fraction", type: "value" },
    ],
  },
  {
    name: "FVSCHEDULE",
    description: "Returns the future value of an initial principal after applying a series of compound interest rates.",
    args: [
      { name: "principal", type: "value" },
      { name: "schedule", type: "range" },
    ],
  },
  {
    name: "IRR",
    description: "Returns the internal rate of return for a series of cash flows.",
    args: [
      { name: "values", type: "range" },
      { name: "guess", type: "value", optional: true },
    ],
  },
  {
    name: "MIRR",
    description: "Returns the modified internal rate of return for a series of periodic cash flows.",
    args: [
      { name: "values", type: "range" },
      { name: "finance_rate", type: "value" },
      { name: "reinvest_rate", type: "value" },
    ],
  },
  {
    name: "ACCRINT",
    description: "Returns the accrued interest for a security that pays periodic interest.",
    args: [
      { name: "issue", type: "value" },
      { name: "first_interest", type: "value" },
      { name: "settlement", type: "value" },
      { name: "rate", type: "value" },
      { name: "par", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "par", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "PRICE",
    description: "Returns the price per $100 face value of a security that pays periodic interest.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "rate", type: "value" },
      { name: "yld", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "pr", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "coupon", type: "value" },
      { name: "yld", type: "value" },
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
      { name: "coupon", type: "value" },
      { name: "yld", type: "value" },
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
      { name: "discount", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "yld", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "YIELDDISC",
    description: "Returns the annual yield for a discounted security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "pr", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "DISC",
    description: "Returns the discount rate for a security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "value" },
      { name: "redemption", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "RECEIVED",
    description: "Returns the amount received at maturity for a fully invested security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "investment", type: "value" },
      { name: "discount", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "INTRATE",
    description: "Returns the interest rate for a fully invested security.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "investment", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "discount", type: "value" },
    ],
  },
  {
    name: "TBILLPRICE",
    description: "Returns the price per $100 face value for a Treasury bill.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "discount", type: "value" },
    ],
  },
  {
    name: "TBILLYIELD",
    description: "Returns the yield for a Treasury bill.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "pr", type: "value" },
    ],
  },
  {
    name: "ODDLPRICE",
    description: "Returns the price per $100 face value of a security with an odd last period.",
    args: [
      { name: "settlement", type: "value" },
      { name: "maturity", type: "value" },
      { name: "last_interest", type: "value" },
      { name: "rate", type: "value" },
      { name: "yld", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "pr", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "yld", type: "value" },
      { name: "redemption", type: "value" },
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
      { name: "rate", type: "value" },
      { name: "pr", type: "value" },
      { name: "redemption", type: "value" },
      { name: "frequency", type: "number" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "AMORDEGRC",
    description: "Returns the depreciation for each accounting period by using a depreciation coefficient.",
    args: [
      { name: "cost", type: "value" },
      { name: "date_purchased", type: "value" },
      { name: "first_period", type: "value" },
      { name: "salvage", type: "value" },
      { name: "period", type: "value" },
      { name: "rate", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "AMORLINC",
    description: "Returns the depreciation for each accounting period.",
    args: [
      { name: "cost", type: "value" },
      { name: "date_purchased", type: "value" },
      { name: "first_period", type: "value" },
      { name: "salvage", type: "value" },
      { name: "period", type: "value" },
      { name: "rate", type: "value" },
      { name: "basis", type: "number", optional: true },
    ],
  },
  {
    name: "SLN",
    description: "Returns the straight-line depreciation of an asset for one period.",
    args: [
      { name: "cost", type: "value" },
      { name: "salvage", type: "value" },
      { name: "life", type: "value" },
    ],
  },
  {
    name: "SYD",
    description: "Returns the sum-of-years' digits depreciation of an asset for a specified period.",
    args: [
      { name: "cost", type: "value" },
      { name: "salvage", type: "value" },
      { name: "life", type: "value" },
      { name: "per", type: "value" },
    ],
  },
  {
    name: "DB",
    description: "Returns the depreciation of an asset for a specified period by using the fixed-declining balance method.",
    args: [
      { name: "cost", type: "value" },
      { name: "salvage", type: "value" },
      { name: "life", type: "value" },
      { name: "period", type: "value" },
      { name: "month", type: "value", optional: true },
    ],
  },
  {
    name: "DDB",
    description: "Returns the depreciation of an asset for a specified period by using the double-declining balance method or another method you specify.",
    args: [
      { name: "cost", type: "value" },
      { name: "salvage", type: "value" },
      { name: "life", type: "value" },
      { name: "period", type: "value" },
      { name: "factor", type: "value", optional: true },
    ],
  },
  {
    name: "VDB",
    description: "Returns the depreciation of an asset for any period you specify, including partial periods.",
    args: [
      { name: "cost", type: "value" },
      { name: "salvage", type: "value" },
      { name: "life", type: "value" },
      { name: "start_period", type: "value" },
      { name: "end_period", type: "value" },
      { name: "factor", type: "value", optional: true },
      { name: "no_switch", type: "boolean", optional: true },
    ],
  },
  {
    name: "XNPV",
    description: "Returns the net present value for a schedule of cash flows that is not necessarily periodic.",
    args: [
      { name: "rate", type: "value" },
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
      { name: "guess", type: "value", optional: true },
    ],
  },
  {
    name: "INDEX",
    description: "Returns the value of an element in a table or an array.",
    args: [
      { name: "array", type: "range" },
      { name: "row_num", type: "value" },
      { name: "column_num", type: "value", optional: true },
      { name: "area_num", type: "number", optional: true },
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
    name: "ISEVEN",
    description: "Returns TRUE if the number is even.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "ISODD",
    description: "Returns TRUE if the number is odd.",
    args: [
      { name: "number", type: "value" },
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
      { name: "num_chars", type: "value", optional: true },
    ],
  },
  {
    name: "LEFTB",
    description: "Returns the leftmost bytes from a text value.",
    args: [
      { name: "text", type: "range" },
      { name: "num_bytes", type: "value", optional: true },
    ],
  },
  {
    name: "RIGHT",
    description: "Returns the rightmost characters from a text value.",
    args: [
      { name: "text", type: "range" },
      { name: "num_chars", type: "value", optional: true },
    ],
  },
  {
    name: "RIGHTB",
    description: "Returns the rightmost bytes from a text value.",
    args: [
      { name: "text", type: "range" },
      { name: "num_bytes", type: "value", optional: true },
    ],
  },
  {
    name: "MID",
    description: "Returns a specific number of characters from a text string, starting at the position you specify.",
    args: [
      { name: "text", type: "range" },
      { name: "start_num", type: "number" },
      { name: "num_chars", type: "value" },
    ],
  },
  {
    name: "MIDB",
    description: "Returns a specific number of bytes from a text string, starting at the position you specify.",
    args: [
      { name: "text", type: "range" },
      { name: "start_num", type: "number" },
      { name: "num_bytes", type: "value" },
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
    name: "LENB",
    description: "Returns the number of bytes used to represent the characters in a text string.",
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
      { name: "num_chars", type: "value" },
      { name: "new_text", type: "value" },
    ],
  },
  {
    name: "REPLACEB",
    description: "Replaces part of a text string with a different text string (byte-based).",
    args: [
      { name: "old_text", type: "range" },
      { name: "start_num", type: "number" },
      { name: "num_bytes", type: "value" },
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
    name: "FINDB",
    description: "Finds one text value within another (case-sensitive, byte-based).",
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
    name: "SEARCHB",
    description: "Finds one text value within another (case-insensitive, byte-based) and returns the position.",
    args: [
      { name: "find_text", type: "value" },
      { name: "within_text", type: "range" },
      { name: "start_num", type: "number", optional: true },
    ],
  },
  {
    name: "EXACT",
    description: "Checks whether two text values are identical.",
    args: [
      { name: "text1", type: "range" },
      { name: "text2", type: "range" },
    ],
  },
  {
    name: "CODE",
    description: "Returns a numeric code for the first character in a text string.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "UNICODE",
    description: "Returns the Unicode (UTF-8) code point for the first character in a text string.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "CHAR",
    description: "Returns the character specified by a number.",
    args: [
      { name: "number", type: "value" },
    ],
  },
  {
    name: "UNICHAR",
    description: "Returns the Unicode character referenced by the given numeric value.",
    args: [
      { name: "number", type: "value" },
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
    name: "DECIMAL",
    description: "Converts a text representation of a number in a given base to decimal.",
    args: [
      { name: "text", type: "range" },
      { name: "radix", type: "number" },
    ],
  },
  {
    name: "BASE",
    description: "Converts a number into a text representation in the given radix (base).",
    args: [
      { name: "number", type: "range" },
      { name: "radix", type: "number" },
      { name: "min_length", type: "value", optional: true },
    ],
  },
  {
    name: "DEC2BIN",
    description: "Converts a decimal number to binary.",
    args: [
      { name: "decimal_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "DEC2HEX",
    description: "Converts a decimal number to hexadecimal.",
    args: [
      { name: "decimal_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "DEC2OCT",
    description: "Converts a decimal number to octal.",
    args: [
      { name: "decimal_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "BIN2DEC",
    description: "Converts a binary number to decimal.",
    args: [
      { name: "binary_number", type: "range" },
    ],
  },
  {
    name: "BIN2HEX",
    description: "Converts a binary number to hexadecimal.",
    args: [
      { name: "binary_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "BIN2OCT",
    description: "Converts a binary number to octal.",
    args: [
      { name: "binary_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "HEX2BIN",
    description: "Converts a hexadecimal number to binary.",
    args: [
      { name: "hex_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "HEX2DEC",
    description: "Converts a hexadecimal number to decimal.",
    args: [
      { name: "hex_number", type: "range" },
    ],
  },
  {
    name: "HEX2OCT",
    description: "Converts a hexadecimal number to octal.",
    args: [
      { name: "hex_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "OCT2BIN",
    description: "Converts an octal number to binary.",
    args: [
      { name: "octal_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "OCT2DEC",
    description: "Converts an octal number to decimal.",
    args: [
      { name: "octal_number", type: "range" },
    ],
  },
  {
    name: "OCT2HEX",
    description: "Converts an octal number to hexadecimal.",
    args: [
      { name: "octal_number", type: "range" },
      { name: "places", type: "value", optional: true },
    ],
  },
  {
    name: "ARABIC",
    description: "Converts a Roman numeral to an Arabic number.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "ROMAN",
    description: "Converts an Arabic numeral to Roman as text.",
    args: [
      { name: "number", type: "range" },
      { name: "form", type: "number", optional: true },
    ],
  },
  {
    name: "CONVERT",
    description: "Converts a number from one measurement system to another.",
    args: [
      { name: "number", type: "value" },
      { name: "from_unit", type: "string" },
      { name: "to_unit", type: "string" },
    ],
  },
  {
    name: "ASC",
    description: "Changes full-width (double-byte) characters to half-width (single-byte) characters.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "DBCS",
    description: "Changes half-width (single-byte) characters to full-width (double-byte) characters.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "PHONETIC",
    description: "Extracts the phonetic (furigana) characters from a text string.",
    args: [
      { name: "reference", type: "range" },
    ],
  },
  {
    name: "ISTHAIDIGIT",
    description: "Returns TRUE if the text is a Thai digit.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "THAIDIGIT",
    description: "Converts Arabic numerals to Thai digits.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "THAISTRINGLENGTH",
    description: "Returns the length of a Thai text string.",
    args: [
      { name: "text", type: "range" },
    ],
  },
  {
    name: "COMPLEX",
    description: "Converts real and imaginary coefficients into a complex number.",
    args: [
      { name: "real_num", type: "value" },
      { name: "i_num", type: "value" },
      { name: "suffix", type: "string", optional: true },
    ],
  },
  {
    name: "IMABS",
    description: "Returns the absolute value (modulus) of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMAGINARY",
    description: "Returns the imaginary coefficient of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMARGUMENT",
    description: "Returns the argument (theta) of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMCONJUGATE",
    description: "Returns the complex conjugate of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMCOS",
    description: "Returns the cosine of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMDIV",
    description: "Returns the quotient of two complex numbers.",
    args: [
      { name: "inumber1", type: "range" },
      { name: "inumber2", type: "range" },
    ],
  },
  {
    name: "IMEXP",
    description: "Returns the exponential of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMLN",
    description: "Returns the natural logarithm of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMLOG10",
    description: "Returns the base-10 logarithm of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMLOG2",
    description: "Returns the base-2 logarithm of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMPOWER",
    description: "Returns a complex number raised to an integer power.",
    args: [
      { name: "inumber", type: "range" },
      { name: "number", type: "value" },
    ],
  },
  {
    name: "IMPRODUCT",
    description: "Returns the product of complex numbers.",
    args: [
      { name: "inumber1", type: "range", repeating: true },
    ],
  },
  {
    name: "IMREAL",
    description: "Returns the real coefficient of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMSIN",
    description: "Returns the sine of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMSQRT",
    description: "Returns the square root of a complex number.",
    args: [
      { name: "inumber", type: "range" },
    ],
  },
  {
    name: "IMSUB",
    description: "Returns the difference of two complex numbers.",
    args: [
      { name: "inumber1", type: "range" },
      { name: "inumber2", type: "range" },
    ],
  },
  {
    name: "IMSUM",
    description: "Returns the sum of complex numbers.",
    args: [
      { name: "inumber1", type: "range", repeating: true },
    ],
  },
  {
    name: "CUBEKPIMEMBER",
    description: "Returns a KPI property and displays the name in the cell.",
    args: [
      { name: "connection", type: "string" },
      { name: "kpi_name", type: "string" },
      { name: "kpi_property", type: "string" },
      { name: "caption", type: "string", optional: true },
    ],
  },
  {
    name: "CUBEMEMBER",
    description: "Returns a member or tuple from the cube.",
    args: [
      { name: "connection", type: "string" },
      { name: "member_expression", type: "string" },
      { name: "caption", type: "string", optional: true },
    ],
  },
  {
    name: "CUBEMEMBERPROPERTY",
    description: "Returns the value of a member property in the cube.",
    args: [
      { name: "connection", type: "string" },
      { name: "member_expression", type: "string" },
      { name: "property", type: "string" },
    ],
  },
  {
    name: "CUBERANKEDMEMBER",
    description: "Returns the nth (ranked) member in a set.",
    args: [
      { name: "connection", type: "string" },
      { name: "set_expression", type: "string" },
      { name: "rank", type: "value" },
      { name: "caption", type: "string", optional: true },
    ],
  },
  {
    name: "CUBESET",
    description: "Defines a calculated set of members or tuples by sending a set expression to the cube.",
    args: [
      { name: "connection", type: "string" },
      { name: "set_expression", type: "string" },
      { name: "caption", type: "string", optional: true },
      { name: "sort_order", type: "value", optional: true },
      { name: "sort_by", type: "string", optional: true },
    ],
  },
  {
    name: "CUBESETCOUNT",
    description: "Returns the number of items in a set.",
    args: [
      { name: "set", type: "string" },
    ],
  },
  {
    name: "CUBEVALUE",
    description: "Returns an aggregated value from the cube.",
    args: [
      { name: "connection", type: "string" },
      { name: "member_expression1", type: "string", repeating: true },
    ],
  },
  {
    name: "RTD",
    description: "Retrieves real-time data from a program that supports COM automation.",
    args: [
      { name: "prog_id", type: "string" },
      { name: "server", type: "string" },
      { name: "topic1", type: "string", repeating: true },
    ],
  },
  {
    name: "CONCAT",
    description: "Combines the text from multiple ranges and/or strings.",
    args: [
      { name: "text1", type: "range", repeating: true },
    ],
  },
  {
    name: "CONCATENATE",
    description: "Combines several text strings into one text string.",
    args: [
      { name: "text1", type: "range", repeating: true },
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
    name: "SEQUENCE",
    description: "Generates a list of sequential numbers in an array.",
    args: [
      { name: "rows", type: "number" },
      { name: "columns", type: "number", optional: true },
      { name: "start", type: "value", optional: true },
      { name: "step", type: "value", optional: true },
    ],
  },
  {
    name: "RANDARRAY",
    description: "Returns an array of random numbers between min and max.",
    args: [
      { name: "rows", type: "number", optional: true },
      { name: "columns", type: "number", optional: true },
      { name: "min", type: "value", optional: true },
      { name: "max", type: "value", optional: true },
      { name: "whole_number", type: "boolean", optional: true },
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
      { name: "row_num", type: "value" },
      { name: "column_num", type: "value" },
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
      { name: "height", type: "value", optional: true },
      { name: "width", type: "value", optional: true },
    ],
  },
  {
    name: "NUMBERVALUE",
    description: "Converts text to a number in a locale-independent way, using custom separators.",
    args: [
      { name: "text", type: "range" },
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
    name: "DOLLAR",
    description: "Converts a number to text using currency format.",
    args: [
      { name: "number", type: "range" },
      { name: "decimals", type: "value", optional: true },
    ],
  },
  {
    name: "BAHTTEXT",
    description: "Converts a number to Thai text.",
    args: [
      { name: "number", type: "range" },
    ],
  },
  {
    name: "FIXED",
    description: "Rounds a number to the specified number of decimals and formats it as text.",
    args: [
      { name: "number", type: "range" },
      { name: "decimals", type: "value", optional: true },
      { name: "no_commas", type: "boolean", optional: true },
    ],
  },
  {
    name: "REPT",
    description: "Repeats text a given number of times.",
    args: [
      { name: "text", type: "range" },
      { name: "number_times", type: "value" },
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
      { name: "value", type: "range" },
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
