import assert from "node:assert/strict";
import test from "node:test";

import { FunctionRegistry } from "../src/functionRegistry.js";

test("FunctionRegistry loads the Rust function catalog (HLOOKUP is present)", () => {
  const registry = new FunctionRegistry();
  assert.ok(registry.getFunction("SEQUENCE"), "Expected SEQUENCE (catalog-only) to be present");
  assert.ok(registry.getFunction("XLOOKUP"), "Expected XLOOKUP to be present");
  assert.ok(registry.getFunction("_xlfn.XLOOKUP"), "Expected _xlfn.XLOOKUP alias to be present");
  assert.ok(registry.isRangeArg("_xlfn.XLOOKUP", 1), "Expected _xlfn.XLOOKUP arg2 to be a range");
  assert.equal(registry.getFunction("SUM")?.minArgs, 0, "Expected SUM minArgs to come from catalog");
  assert.equal(registry.getFunction("SUM")?.maxArgs, 255, "Expected SUM maxArgs to come from catalog");
  assert.equal(registry.getArgType("PV", 0), "number", "Expected PV arg1 type to come from catalog arg_types");
  assert.ok(
    registry.getFunction("HLOOKUP"),
    `Expected HLOOKUP to be present, got: ${registry.list().map(f => f.name).join(", ")}`
  );
});

test("FunctionRegistry falls back to curated defaults when catalog is missing/invalid", () => {
  const missingCatalog = new FunctionRegistry(undefined, { catalog: null });
  assert.ok(missingCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    missingCatalog.getFunction("SEQUENCE"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is missing"
  );

  const invalidCatalog = new FunctionRegistry(undefined, { catalog: { functions: [{ nope: true }] } });
  assert.ok(invalidCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    invalidCatalog.getFunction("SEQUENCE"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is invalid"
  );
});

test("FunctionRegistry uses curated range metadata for common multi-range functions", () => {
  const registry = new FunctionRegistry();

  // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  assert.ok(registry.isRangeArg("SUMIFS", 0), "Expected SUMIFS arg1 to be a range");
  assert.ok(registry.isRangeArg("SUMIFS", 1), "Expected SUMIFS arg2 to be a range");
  assert.equal(registry.isRangeArg("SUMIFS", 2), false, "Expected SUMIFS arg3 to be a value");
  assert.ok(registry.isRangeArg("SUMIFS", 3), "Expected SUMIFS arg4 (criteria_range2) to be a range");
  assert.equal(registry.isRangeArg("SUMIFS", 4), false, "Expected SUMIFS arg5 (criteria2) to be a value");

  // _xlfn aliases should preserve the curated signatures.
  assert.ok(registry.isRangeArg("_xlfn.SUMIFS", 0), "Expected _xlfn.SUMIFS arg1 to be a range");
  assert.ok(registry.isRangeArg("_xlfn.FILTER", 0), "Expected _xlfn.FILTER arg1 to be a range");
  assert.ok(registry.isRangeArg("_xlfn.FILTER", 1), "Expected _xlfn.FILTER arg2 to be a range");

  // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
  assert.equal(registry.isRangeArg("TEXTJOIN", 0), false, "Expected TEXTJOIN delimiter not to be a range");
  assert.equal(registry.isRangeArg("TEXTJOIN", 1), false, "Expected TEXTJOIN ignore_empty not to be a range");
  assert.ok(registry.isRangeArg("TEXTJOIN", 2), "Expected TEXTJOIN text1 to be a range");
  assert.ok(registry.isRangeArg("TEXTJOIN", 3), "Expected TEXTJOIN text2 to be a range (varargs)");

  // SUBTOTAL(function_num, ref1, [ref2], ...)
  assert.equal(registry.isRangeArg("SUBTOTAL", 0), false, "Expected SUBTOTAL function_num not to be a range");
  assert.ok(registry.isRangeArg("SUBTOTAL", 1), "Expected SUBTOTAL ref1 to be a range");
  assert.ok(registry.isRangeArg("SUBTOTAL", 2), "Expected SUBTOTAL ref2 to be a range (varargs)");

  // AGGREGATE(function_num, options, ref1, [ref2], ...)
  assert.equal(registry.isRangeArg("AGGREGATE", 0), false, "Expected AGGREGATE function_num not to be a range");
  assert.equal(registry.isRangeArg("AGGREGATE", 1), false, "Expected AGGREGATE options not to be a range");
  assert.ok(registry.isRangeArg("AGGREGATE", 2), "Expected AGGREGATE ref1 to be a range");
  assert.ok(registry.isRangeArg("AGGREGATE", 3), "Expected AGGREGATE ref2 to be a range (varargs)");

  // FORECAST.LINEAR(x, known_y's, known_x's)
  assert.equal(
    registry.isRangeArg("FORECAST.LINEAR", 0),
    false,
    "Expected FORECAST.LINEAR x not to be a range"
  );
  assert.ok(registry.isRangeArg("FORECAST.LINEAR", 1), "Expected FORECAST.LINEAR known_ys to be a range");
  assert.ok(registry.isRangeArg("FORECAST.LINEAR", 2), "Expected FORECAST.LINEAR known_xs to be a range");

  // HSTACK/VSTACK(array1, [array2], ...)
  assert.ok(registry.isRangeArg("HSTACK", 0), "Expected HSTACK array1 to be a range");
  assert.ok(registry.isRangeArg("HSTACK", 1), "Expected HSTACK array2 to be a range (varargs)");
  assert.ok(registry.isRangeArg("_xlfn.HSTACK", 0), "Expected _xlfn.HSTACK array1 to be a range");

  // Dot-name functions should also work with _xlfn aliases.
  assert.ok(
    registry.isRangeArg("_xlfn.FORECAST.LINEAR", 1),
    "Expected _xlfn.FORECAST.LINEAR known_ys to be a range"
  );

  // OFFSET(reference, rows, cols, ...)
  assert.ok(registry.isRangeArg("OFFSET", 0), "Expected OFFSET reference to be a range");
  assert.ok(registry.isRangeArg("_xlfn.OFFSET", 0), "Expected _xlfn.OFFSET reference to be a range");

  // Common stats varargs: STDEV.S / VAR.S
  assert.ok(registry.isRangeArg("STDEV.S", 0), "Expected STDEV.S arg1 to be a range");
  assert.ok(registry.isRangeArg("VAR.S", 0), "Expected VAR.S arg1 to be a range");

  // Matrix functions
  assert.ok(registry.isRangeArg("MMULT", 0), "Expected MMULT array1 to be a range");
  assert.ok(registry.isRangeArg("MMULT", 1), "Expected MMULT array2 to be a range");
  assert.ok(registry.isRangeArg("MDETERM", 0), "Expected MDETERM array to be a range");
  assert.ok(registry.isRangeArg("MINVERSE", 0), "Expected MINVERSE array to be a range");

  // Statistical test functions with dot/legacy names
  assert.ok(registry.isRangeArg("T.TEST", 0), "Expected T.TEST array1 to be a range");
  assert.ok(registry.isRangeArg("TTEST", 0), "Expected TTEST array1 to be a range");
  assert.ok(registry.isRangeArg("F.TEST", 0), "Expected F.TEST array1 to be a range");
  assert.ok(registry.isRangeArg("FTEST", 0), "Expected FTEST array1 to be a range");
  assert.ok(registry.isRangeArg("Z.TEST", 0), "Expected Z.TEST array to be a range");
  assert.ok(registry.isRangeArg("ZTEST", 0), "Expected ZTEST array to be a range");

  // Additional common stats functions
  assert.ok(registry.isRangeArg("PERCENTILE.EXC", 0), "Expected PERCENTILE.EXC array to be a range");
  assert.ok(registry.isRangeArg("QUARTILE.EXC", 0), "Expected QUARTILE.EXC array to be a range");
  assert.ok(registry.isRangeArg("RANK.AVG", 1), "Expected RANK.AVG ref to be a range");
  assert.ok(registry.isRangeArg("MODE.SNGL", 0), "Expected MODE.SNGL arg1 to be a range");
  assert.ok(registry.isRangeArg("TRIMMEAN", 0), "Expected TRIMMEAN array to be a range");

  // Dynamic array helpers
  assert.ok(registry.isRangeArg("BYROW", 0), "Expected BYROW array to be a range");
  assert.ok(registry.isRangeArg("BYCOL", 0), "Expected BYCOL array to be a range");
  assert.equal(registry.getFunction("MAKEARRAY")?.args?.[2]?.name, "lambda", "Expected MAKEARRAY arg3 to be lambda");
  assert.ok(registry.isRangeArg("MAP", 0), "Expected MAP array to be a range");
  assert.equal(registry.isRangeArg("MAP", 1), false, "Expected MAP lambda not to be a range");

  // REDUCE/SCAN support both 2-arg and 3-arg forms, so we treat the leading arg as range-like.
  assert.ok(registry.isRangeArg("REDUCE", 0), "Expected REDUCE arg1 to be a range");
  assert.ok(registry.isRangeArg("REDUCE", 1), "Expected REDUCE arg2 (array) to be a range");
  assert.equal(registry.isRangeArg("REDUCE", 2), false, "Expected REDUCE lambda not to be a range");

  assert.ok(registry.isRangeArg("SCAN", 0), "Expected SCAN arg1 to be a range");
  assert.ok(registry.isRangeArg("SCAN", 1), "Expected SCAN arg2 (array) to be a range");
  assert.equal(registry.isRangeArg("SCAN", 2), false, "Expected SCAN lambda not to be a range");

  assert.ok(registry.isRangeArg("_xlfn.MAP", 0), "Expected _xlfn.MAP array to be a range");

  // Conditional logic with repeating (test/value) pairs
  const ifs = registry.getFunction("IFS");
  assert.ok(ifs, "Expected IFS to have a curated signature");
  assert.equal(ifs?.args?.[0]?.name, "logical_test1", "Expected IFS arg1 name to be logical_test1");
  assert.ok(ifs?.args?.[0]?.repeating, "Expected IFS logical_test1 to mark a repeating group");
  assert.equal(ifs?.args?.[1]?.name, "value_if_true1", "Expected IFS arg2 name to be value_if_true1");

  const sw = registry.getFunction("SWITCH");
  assert.ok(sw, "Expected SWITCH to have a curated signature");
  assert.equal(sw?.args?.[0]?.name, "expression", "Expected SWITCH arg1 name to be expression");
  assert.ok(sw?.args?.[1]?.repeating, "Expected SWITCH value1 to mark a repeating group");
  assert.ok(registry.getFunction("IFERROR"), "Expected IFERROR to have a curated signature");
  assert.ok(registry.getFunction("IFNA"), "Expected IFNA to have a curated signature");
  assert.equal(registry.getArgType("LET", 0), "string", "Expected LET name1 to be a string-like identifier");
  assert.equal(registry.getArgType("LET", 2), "value", "Expected LET calculation to be a value");
  assert.equal(registry.getArgType("LAMBDA", 0), "string", "Expected LAMBDA parameter1 to be string-like");
  assert.equal(registry.getArgType("LAMBDA", 1), "value", "Expected LAMBDA calculation to be a value");

  // Legacy descriptive stats
  assert.ok(registry.isRangeArg("PERCENTILE", 0), "Expected PERCENTILE array to be a range");
  assert.ok(registry.isRangeArg("QUARTILE", 0), "Expected QUARTILE array to be a range");
  assert.ok(registry.isRangeArg("RANK", 1), "Expected RANK ref to be a range");
  assert.ok(registry.isRangeArg("PERCENTRANK", 0), "Expected PERCENTRANK array to be a range");
  assert.ok(registry.isRangeArg("PERCENTRANK.INC", 0), "Expected PERCENTRANK.INC array to be a range");
  assert.ok(registry.isRangeArg("PERCENTRANK.EXC", 0), "Expected PERCENTRANK.EXC array to be a range");

  // Legacy varargs
  assert.ok(registry.isRangeArg("STDEV", 0), "Expected STDEV arg1 to be a range");
  assert.ok(registry.isRangeArg("VAR", 0), "Expected VAR arg1 to be a range");

  // Database functions (database + criteria are ranges)
  assert.ok(registry.isRangeArg("DSUM", 0), "Expected DSUM database to be a range");
  assert.ok(registry.isRangeArg("DSUM", 2), "Expected DSUM criteria to be a range");
  assert.ok(registry.isRangeArg("DCOUNT", 0), "Expected DCOUNT database to be a range");
  assert.ok(registry.isRangeArg("DCOUNT", 2), "Expected DCOUNT criteria to be a range");

  // Size helpers
  assert.ok(registry.isRangeArg("ROWS", 0), "Expected ROWS array to be a range");
  assert.ok(registry.isRangeArg("COLUMNS", 0), "Expected COLUMNS array to be a range");
  assert.ok(registry.isRangeArg("ROW", 0), "Expected ROW reference to be a range");
  assert.ok(registry.isRangeArg("COLUMN", 0), "Expected COLUMN reference to be a range");

  // Reference/info helpers
  assert.ok(registry.isRangeArg("AREAS", 0), "Expected AREAS reference to be a range");
  assert.equal(registry.isRangeArg("CELL", 0), false, "Expected CELL info_type not to be a range");
  assert.ok(registry.isRangeArg("CELL", 1), "Expected CELL reference to be a range");
  assert.ok(registry.isRangeArg("FORMULATEXT", 0), "Expected FORMULATEXT reference to be a range");
  assert.ok(registry.isRangeArg("ISFORMULA", 0), "Expected ISFORMULA reference to be a range");
  assert.ok(registry.isRangeArg("SHEET", 0), "Expected SHEET value to be a range");
  assert.ok(registry.isRangeArg("SHEETS", 0), "Expected SHEETS reference to be a range");
  assert.ok(registry.isRangeArg("GETPIVOTDATA", 1), "Expected GETPIVOTDATA pivot_table to be a range");
  assert.ok(
    registry.getFunction("GETPIVOTDATA")?.args?.[2]?.repeating,
    "Expected GETPIVOTDATA field/item pairs to repeat"
  );

  // Finance/stat functions that take ranges (catalog arg_types are too coarse)
  assert.equal(registry.isRangeArg("FVSCHEDULE", 0), false, "Expected FVSCHEDULE principal not to be a range");
  assert.ok(registry.isRangeArg("FVSCHEDULE", 1), "Expected FVSCHEDULE schedule to be a range");
  assert.ok(registry.isRangeArg("MIRR", 0), "Expected MIRR values to be a range");
  assert.ok(registry.isRangeArg("PROB", 0), "Expected PROB x_range to be a range");
  assert.ok(registry.isRangeArg("PROB", 1), "Expected PROB prob_range to be a range");
  assert.ok(registry.isRangeArg("SERIESSUM", 3), "Expected SERIESSUM coefficients to be a range");

  // Time-series forecasting functions
  assert.ok(registry.isRangeArg("FORECAST.ETS", 1), "Expected FORECAST.ETS values to be a range");
  assert.ok(registry.isRangeArg("FORECAST.ETS", 2), "Expected FORECAST.ETS timeline to be a range");

  // Date functions with optional holiday ranges
  assert.equal(registry.isRangeArg("WORKDAY", 0), false, "Expected WORKDAY start_date not to be a range");
  assert.equal(registry.isRangeArg("WORKDAY", 1), false, "Expected WORKDAY days not to be a range");
  assert.ok(registry.isRangeArg("WORKDAY", 2), "Expected WORKDAY holidays to be a range");
  assert.ok(registry.isRangeArg("WORKDAY.INTL", 3), "Expected WORKDAY.INTL holidays to be a range");
  assert.ok(registry.isRangeArg("NETWORKDAYS", 2), "Expected NETWORKDAYS holidays to be a range");
  assert.ok(registry.isRangeArg("NETWORKDAYS.INTL", 3), "Expected NETWORKDAYS.INTL holidays to be a range");

  // String-only parameters where suggesting cell refs is undesirable (argument hinting)
  assert.equal(registry.getArgType("DATEDIF", 2), "string", "Expected DATEDIF unit to be a string");
  assert.equal(registry.getArgType("TEXT", 1), "string", "Expected TEXT format_text to be a string");
  assert.equal(registry.getArgType("HYPERLINK", 0), "string", "Expected HYPERLINK link_location to be a string");
  assert.ok(
    registry.getFunction("HYPERLINK")?.args?.[1]?.optional,
    "Expected HYPERLINK friendly_name to be optional"
  );
  assert.equal(registry.getArgType("INFO", 0), "string", "Expected INFO type_text to be a string");
  assert.equal(registry.getArgType("INDIRECT", 0), "string", "Expected INDIRECT ref_text to be a string");
  assert.equal(registry.getArgType("ADDRESS", 4), "string", "Expected ADDRESS sheet_text to be a string");
  assert.equal(registry.getArgType("NUMBERVALUE", 1), "string", "Expected NUMBERVALUE decimal_separator to be a string");
  assert.equal(registry.getArgType("NUMBERVALUE", 2), "string", "Expected NUMBERVALUE group_separator to be a string");
  assert.equal(registry.getArgType("IMAGE", 0), "string", "Expected IMAGE source to be a string");

  // Date/time helpers with more descriptive arg naming
  assert.equal(registry.getFunction("DAYS")?.args?.[0]?.name, "end_date", "Expected DAYS arg1 to be end_date");
  assert.equal(registry.getFunction("DAYS360")?.args?.[2]?.name, "method", "Expected DAYS360 arg3 to be method");
  assert.equal(registry.getArgType("DAYS360", 2), "boolean", "Expected DAYS360 method to be boolean");
  assert.equal(registry.getFunction("YEARFRAC")?.args?.[2]?.name, "basis", "Expected YEARFRAC arg3 to be basis");
  assert.equal(registry.getArgType("YEARFRAC", 2), "number", "Expected YEARFRAC basis to be number");
  assert.equal(
    registry.getFunction("DATEVALUE")?.args?.[0]?.name,
    "date_text",
    "Expected DATEVALUE arg1 to be date_text"
  );

  // Bond/treasury functions: ensure arg naming matches enum indices in TabCompletionEngine.
  const price = registry.getFunction("PRICE");
  assert.ok(price, "Expected PRICE to have a curated signature");
  assert.equal(price?.args?.[0]?.name, "settlement", "Expected PRICE arg1 to be settlement");
  assert.equal(price?.args?.[5]?.name, "frequency", "Expected PRICE arg6 to be frequency");
  assert.ok(price?.args?.[6]?.optional, "Expected PRICE basis to be optional");

  assert.equal(registry.getArgType("ACCRINT", 7), "boolean", "Expected ACCRINT calc_method to be boolean");
  assert.equal(registry.getFunction("COUPDAYBS")?.args?.[2]?.name, "frequency", "Expected COUPDAYBS arg3 to be frequency");
  assert.equal(registry.getFunction("TBILLYIELD")?.args?.[2]?.name, "pr", "Expected TBILLYIELD arg3 to be pr");
});
