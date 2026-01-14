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
  assert.equal(
    registry.getArgType("RANDBETWEEN", 0),
    "number",
    "Expected RANDBETWEEN arg1 type to come from catalog arg_types"
  );
  assert.ok(
    registry.getFunction("HLOOKUP"),
    `Expected HLOOKUP to be present, got: ${registry.list().map(f => f.name).join(", ")}`
  );
});

test("FunctionRegistry falls back to curated defaults when catalog is missing/invalid", () => {
  const missingCatalog = new FunctionRegistry(undefined, { catalog: null });
  assert.ok(missingCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    missingCatalog.getFunction("ACOTH"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is missing"
  );

  const invalidCatalog = new FunctionRegistry(undefined, { catalog: { functions: [{ nope: true }] } });
  assert.ok(invalidCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    invalidCatalog.getFunction("ACOTH"),
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
  assert.ok(registry.isRangeArg("CONCAT", 0), "Expected CONCAT text1 to be a range");
  assert.ok(registry.isRangeArg("CONCATENATE", 0), "Expected CONCATENATE text1 to be a range");
  assert.ok(registry.isRangeArg("REPT", 0), "Expected REPT text to be a range");
  assert.equal(registry.getFunction("REPT")?.args?.[1]?.name, "number_times", "Expected REPT arg2 to be number_times");
  assert.equal(registry.getArgType("REPT", 1), "value", "Expected REPT number_times to be value-like");
  assert.ok(registry.isRangeArg("DOLLAR", 0), "Expected DOLLAR number to be a range");
  assert.ok(registry.getFunction("DOLLAR")?.args?.[1]?.optional, "Expected DOLLAR decimals to be optional");
  assert.equal(registry.getArgType("DOLLAR", 1), "value", "Expected DOLLAR decimals to be value-like");
  assert.ok(registry.isRangeArg("BAHTTEXT", 0), "Expected BAHTTEXT number to be a range");
  assert.ok(registry.isRangeArg("FIXED", 0), "Expected FIXED number to be a range");
  assert.equal(registry.getArgType("FIXED", 1), "value", "Expected FIXED decimals to be value-like");
  assert.equal(registry.getArgType("FIXED", 2), "boolean", "Expected FIXED no_commas to be boolean");

  // Common scalar math functions should treat numeric inputs as value-like (often cell-referenced).
  assert.equal(registry.getArgType("ABS", 0), "value", "Expected ABS number to be value-like");
  assert.equal(registry.getFunction("ROUND")?.args?.[1]?.name, "num_digits", "Expected ROUND arg2 to be num_digits");
  assert.equal(registry.getArgType("ROUND", 0), "value", "Expected ROUND number to be value-like");
  assert.equal(registry.getArgType("ROUND", 1), "value", "Expected ROUND num_digits to be value-like");
  assert.equal(registry.getArgType("ROUNDUP", 0), "value", "Expected ROUNDUP number to be value-like");
  assert.equal(registry.getArgType("ROUNDUP", 1), "value", "Expected ROUNDUP num_digits to be value-like");
  assert.equal(registry.getArgType("ROUNDDOWN", 0), "value", "Expected ROUNDDOWN number to be value-like");
  assert.equal(registry.getArgType("ROUNDDOWN", 1), "value", "Expected ROUNDDOWN num_digits to be value-like");
  assert.equal(registry.getArgType("INT", 0), "value", "Expected INT number to be value-like");
  assert.ok(registry.getFunction("TRUNC")?.args?.[1]?.optional, "Expected TRUNC num_digits to be optional");
  assert.equal(registry.getArgType("TRUNC", 1), "value", "Expected TRUNC num_digits to be value-like");
  assert.equal(registry.getArgType("POWER", 0), "value", "Expected POWER number to be value-like");
  assert.equal(registry.getArgType("POWER", 1), "value", "Expected POWER power to be value-like");
  assert.equal(registry.getArgType("SQRT", 0), "value", "Expected SQRT number to be value-like");
  assert.equal(registry.getArgType("LOG", 0), "value", "Expected LOG number to be value-like");
  assert.ok(registry.getFunction("LOG")?.args?.[1]?.optional, "Expected LOG base to be optional");
  assert.equal(registry.getArgType("SIN", 0), "value", "Expected SIN number to be value-like");
  assert.equal(registry.getArgType("BITAND", 0), "value", "Expected BITAND number1 to be value-like");
  assert.equal(registry.getArgType("BITAND", 1), "value", "Expected BITAND number2 to be value-like");
  assert.equal(registry.getArgType("BITOR", 0), "value", "Expected BITOR number1 to be value-like");
  assert.equal(registry.getArgType("BITOR", 1), "value", "Expected BITOR number2 to be value-like");
  assert.equal(registry.getArgType("BITXOR", 0), "value", "Expected BITXOR number1 to be value-like");
  assert.equal(registry.getArgType("BITXOR", 1), "value", "Expected BITXOR number2 to be value-like");
  assert.equal(registry.getArgType("BITLSHIFT", 0), "value", "Expected BITLSHIFT number to be value-like");
  assert.equal(registry.getArgType("BITLSHIFT", 1), "value", "Expected BITLSHIFT shift_amount to be value-like");
  assert.equal(registry.getArgType("BITRSHIFT", 0), "value", "Expected BITRSHIFT number to be value-like");
  assert.equal(registry.getArgType("BITRSHIFT", 1), "value", "Expected BITRSHIFT shift_amount to be value-like");
  assert.equal(registry.getArgType("ISEVEN", 0), "value", "Expected ISEVEN number to be value-like");
  assert.equal(registry.getArgType("ISODD", 0), "value", "Expected ISODD number to be value-like");
  assert.equal(registry.getArgType("FACT", 0), "value", "Expected FACT number to be value-like");
  assert.equal(registry.getArgType("FACTDOUBLE", 0), "value", "Expected FACTDOUBLE number to be value-like");
  assert.equal(registry.getArgType("COMBIN", 0), "value", "Expected COMBIN number to be value-like");
  assert.equal(registry.getArgType("COMBIN", 1), "value", "Expected COMBIN number_chosen to be value-like");
  assert.equal(registry.getArgType("COMBINA", 0), "value", "Expected COMBINA number to be value-like");
  assert.equal(registry.getArgType("PERMUT", 0), "value", "Expected PERMUT number to be value-like");
  assert.equal(registry.getArgType("PERMUTATIONA", 0), "value", "Expected PERMUTATIONA number to be value-like");
  assert.equal(registry.getArgType("GAMMA", 0), "value", "Expected GAMMA number to be value-like");
  assert.equal(registry.getArgType("GAMMALN", 0), "value", "Expected GAMMALN x to be value-like");
  assert.equal(registry.getArgType("GAMMALN.PRECISE", 0), "value", "Expected GAMMALN.PRECISE x to be value-like");
  assert.equal(registry.getArgType("GAUSS", 0), "value", "Expected GAUSS z to be value-like");
  assert.equal(registry.getArgType("ERF", 0), "value", "Expected ERF lower_limit to be value-like");
  assert.ok(registry.getFunction("ERF")?.args?.[1]?.optional, "Expected ERF upper_limit to be optional");
  assert.equal(registry.getArgType("ERFC", 0), "value", "Expected ERFC x to be value-like");
  assert.equal(registry.getArgType("BESSELI", 0), "value", "Expected BESSELI x to be value-like");
  assert.equal(registry.getArgType("BESSELI", 1), "number", "Expected BESSELI n to be numeric");
  assert.equal(registry.getArgType("BESSELJ", 0), "value", "Expected BESSELJ x to be value-like");
  assert.equal(registry.getArgType("BESSELK", 0), "value", "Expected BESSELK x to be value-like");
  assert.equal(registry.getArgType("BESSELY", 0), "value", "Expected BESSELY x to be value-like");
  assert.equal(registry.getArgType("DELTA", 0), "value", "Expected DELTA number1 to be value-like");
  assert.equal(registry.getArgType("DELTA", 1), "number", "Expected DELTA number2 to be numeric");
  assert.ok(registry.getFunction("DELTA")?.args?.[1]?.optional, "Expected DELTA number2 to be optional");
  assert.equal(registry.getArgType("GESTEP", 0), "value", "Expected GESTEP number to be value-like");
  assert.equal(registry.getArgType("GESTEP", 1), "number", "Expected GESTEP step to be numeric");
  assert.ok(registry.getFunction("GESTEP")?.args?.[1]?.optional, "Expected GESTEP step to be optional");
  assert.equal(registry.getArgType("COSH", 0), "value", "Expected COSH number to be value-like");
  assert.equal(registry.getArgType("SINH", 0), "value", "Expected SINH number to be value-like");
  assert.equal(registry.getArgType("TANH", 0), "value", "Expected TANH number to be value-like");
  assert.equal(registry.getArgType("CSC", 0), "value", "Expected CSC number to be value-like");
  assert.equal(registry.getArgType("SEC", 0), "value", "Expected SEC number to be value-like");
  assert.equal(registry.getArgType("COT", 0), "value", "Expected COT number to be value-like");
  assert.equal(registry.getArgType("CSCH", 0), "value", "Expected CSCH number to be value-like");
  assert.equal(registry.getArgType("SECH", 0), "value", "Expected SECH number to be value-like");
  assert.equal(registry.getArgType("COTH", 0), "value", "Expected COTH number to be value-like");
  assert.equal(registry.getArgType("ASINH", 0), "value", "Expected ASINH number to be value-like");
  assert.equal(registry.getArgType("ACOSH", 0), "value", "Expected ACOSH number to be value-like");
  assert.equal(registry.getArgType("ATANH", 0), "value", "Expected ATANH number to be value-like");
  assert.equal(registry.getArgType("ACOT", 0), "value", "Expected ACOT number to be value-like");

  // TEXTAFTER/TEXTBEFORE are curated (not present in the Rust catalog yet).
  assert.ok(registry.getFunction("TEXTAFTER"), "Expected TEXTAFTER to be present");
  assert.equal(registry.getFunction("TEXTAFTER")?.minArgs, 2, "Expected TEXTAFTER minArgs to be 2");
  assert.ok(registry.isRangeArg("TEXTAFTER", 0), "Expected TEXTAFTER text to be a range");
  assert.equal(registry.getArgType("TEXTAFTER", 1), "string", "Expected TEXTAFTER delimiter to be string-like");
  assert.ok(registry.getFunction("_xlfn.TEXTAFTER"), "Expected _xlfn.TEXTAFTER alias to be present");
  assert.ok(registry.getFunction("TEXTBEFORE"), "Expected TEXTBEFORE to be present");

  // Common text helpers often take ranges/arrays, so we mark their text arguments as range-like.
  assert.ok(registry.isRangeArg("LEFT", 0), "Expected LEFT text to be a range");
  assert.ok(registry.getFunction("LEFT")?.args?.[1]?.optional, "Expected LEFT num_chars to be optional");
  assert.equal(registry.getArgType("LEFT", 1), "value", "Expected LEFT num_chars to be value-like");
  assert.ok(registry.isRangeArg("LEFTB", 0), "Expected LEFTB text to be a range");
  assert.equal(registry.getArgType("LEFTB", 1), "value", "Expected LEFTB num_bytes to be value-like");
  assert.ok(registry.isRangeArg("RIGHT", 0), "Expected RIGHT text to be a range");
  assert.equal(registry.getArgType("RIGHT", 1), "value", "Expected RIGHT num_chars to be value-like");
  assert.ok(registry.isRangeArg("RIGHTB", 0), "Expected RIGHTB text to be a range");
  assert.equal(registry.getArgType("RIGHTB", 1), "value", "Expected RIGHTB num_bytes to be value-like");
  assert.ok(registry.isRangeArg("MID", 0), "Expected MID text to be a range");
  assert.equal(registry.getArgType("MID", 2), "value", "Expected MID num_chars to be value-like");
  assert.equal(registry.getArgType("MIDB", 2), "value", "Expected MIDB num_bytes to be value-like");
  assert.ok(registry.isRangeArg("LEN", 0), "Expected LEN text to be a range");
  assert.ok(registry.isRangeArg("LENB", 0), "Expected LENB text to be a range");
  assert.ok(registry.isRangeArg("TRIM", 0), "Expected TRIM text to be a range");
  assert.ok(registry.isRangeArg("SUBSTITUTE", 0), "Expected SUBSTITUTE text to be a range");
  assert.ok(registry.isRangeArg("FIND", 1), "Expected FIND within_text to be a range");
  assert.ok(registry.isRangeArg("FINDB", 1), "Expected FINDB within_text to be a range");
  assert.equal(registry.getArgType("REPLACE", 2), "value", "Expected REPLACE num_chars to be value-like");
  assert.equal(registry.getArgType("REPLACEB", 2), "value", "Expected REPLACEB num_bytes to be value-like");
  assert.ok(registry.isRangeArg("EXACT", 0), "Expected EXACT text1 to be a range");
  assert.ok(registry.isRangeArg("CODE", 0), "Expected CODE text to be a range");
  assert.ok(registry.isRangeArg("UNICODE", 0), "Expected UNICODE text to be a range");
  assert.equal(registry.getArgType("CHAR", 0), "value", "Expected CHAR number to be value-like");
  assert.equal(registry.getArgType("UNICHAR", 0), "value", "Expected UNICHAR number to be value-like");
  assert.ok(registry.isRangeArg("VALUE", 0), "Expected VALUE text to be a range");
  assert.ok(registry.isRangeArg("DECIMAL", 0), "Expected DECIMAL text to be a range");
  assert.ok(registry.isRangeArg("BASE", 0), "Expected BASE number to be a range");
  assert.equal(registry.getArgType("BASE", 1), "number", "Expected BASE radix to be a number");
  assert.ok(registry.getFunction("BASE")?.args?.[2]?.optional, "Expected BASE min_length to be optional");
  assert.equal(registry.getArgType("BASE", 2), "value", "Expected BASE min_length to be value-like");
  assert.ok(registry.isRangeArg("DEC2BIN", 0), "Expected DEC2BIN decimal_number to be a range");
  assert.ok(registry.getFunction("DEC2BIN")?.args?.[1]?.optional, "Expected DEC2BIN places to be optional");
  assert.equal(registry.getArgType("DEC2BIN", 1), "value", "Expected DEC2BIN places to be value-like");
  assert.ok(registry.isRangeArg("DEC2HEX", 0), "Expected DEC2HEX decimal_number to be a range");
  assert.equal(registry.getArgType("DEC2HEX", 1), "value", "Expected DEC2HEX places to be value-like");
  assert.ok(registry.isRangeArg("DEC2OCT", 0), "Expected DEC2OCT decimal_number to be a range");
  assert.equal(registry.getArgType("DEC2OCT", 1), "value", "Expected DEC2OCT places to be value-like");
  assert.equal(registry.getArgType("CONVERT", 0), "value", "Expected CONVERT number to be value-like");
  assert.equal(registry.getArgType("CONVERT", 1), "string", "Expected CONVERT from_unit to be string-like");
  assert.ok(registry.isRangeArg("BIN2DEC", 0), "Expected BIN2DEC binary_number to be a range");
  assert.ok(registry.isRangeArg("HEX2DEC", 0), "Expected HEX2DEC hex_number to be a range");
  assert.ok(registry.isRangeArg("OCT2DEC", 0), "Expected OCT2DEC octal_number to be a range");
  assert.ok(registry.getFunction("BIN2HEX")?.args?.[1]?.optional, "Expected BIN2HEX places to be optional");
  assert.equal(registry.getArgType("BIN2HEX", 1), "value", "Expected BIN2HEX places to be value-like");
  assert.ok(registry.getFunction("HEX2BIN")?.args?.[1]?.optional, "Expected HEX2BIN places to be optional");
  assert.equal(registry.getArgType("HEX2BIN", 1), "value", "Expected HEX2BIN places to be value-like");
  assert.equal(registry.getArgType("HEX2OCT", 1), "value", "Expected HEX2OCT places to be value-like");
  assert.equal(registry.getArgType("OCT2BIN", 1), "value", "Expected OCT2BIN places to be value-like");
  assert.equal(registry.getArgType("OCT2HEX", 1), "value", "Expected OCT2HEX places to be value-like");
  assert.ok(registry.isRangeArg("ARABIC", 0), "Expected ARABIC text to be a range");
  assert.ok(registry.isRangeArg("ROMAN", 0), "Expected ROMAN number to be a range");
  assert.ok(registry.getFunction("ROMAN")?.args?.[1]?.optional, "Expected ROMAN form to be optional");
  assert.ok(registry.isRangeArg("ASC", 0), "Expected ASC text to be a range");
  assert.ok(registry.isRangeArg("PHONETIC", 0), "Expected PHONETIC reference to be a range");
  assert.equal(registry.getArgType("COMPLEX", 0), "value", "Expected COMPLEX real_num to be value-like");
  assert.equal(registry.getArgType("COMPLEX", 1), "value", "Expected COMPLEX i_num to be value-like");
  assert.equal(registry.getArgType("COMPLEX", 2), "string", "Expected COMPLEX suffix to be string-like");
  assert.ok(registry.isRangeArg("IMABS", 0), "Expected IMABS inumber to be a range");
  assert.ok(registry.isRangeArg("IMDIV", 1), "Expected IMDIV inumber2 to be a range");
  assert.ok(registry.getFunction("IMSUM")?.args?.[0]?.repeating, "Expected IMSUM to accept repeating inumber args");
  assert.equal(registry.getArgType("RTD", 0), "string", "Expected RTD prog_id to be string-like");
  assert.ok(registry.getFunction("CUBEVALUE")?.args?.[1]?.repeating, "Expected CUBEVALUE member args to repeat");
  assert.equal(registry.getArgType("CUBERANKEDMEMBER", 2), "value", "Expected CUBERANKEDMEMBER rank to be value-like");
  assert.equal(registry.getArgType("CUBESET", 3), "value", "Expected CUBESET sort_order to be value-like");

  // SUBTOTAL(function_num, ref1, [ref2], ...)
  assert.equal(registry.isRangeArg("SUBTOTAL", 0), false, "Expected SUBTOTAL function_num not to be a range");
  assert.ok(registry.isRangeArg("SUBTOTAL", 1), "Expected SUBTOTAL ref1 to be a range");
  assert.ok(registry.isRangeArg("SUBTOTAL", 2), "Expected SUBTOTAL ref2 to be a range (varargs)");

  // Math varargs often operate over ranges.
  assert.ok(registry.isRangeArg("GCD", 0), "Expected GCD number1 to be a range");
  assert.ok(registry.isRangeArg("LCM", 0), "Expected LCM number1 to be a range");
  assert.ok(registry.isRangeArg("MULTINOMIAL", 0), "Expected MULTINOMIAL number1 to be a range");

  // AGGREGATE(function_num, options, ref1, [ref2], ...)
  assert.equal(registry.isRangeArg("AGGREGATE", 0), false, "Expected AGGREGATE function_num not to be a range");
  assert.equal(registry.isRangeArg("AGGREGATE", 1), false, "Expected AGGREGATE options not to be a range");
  assert.ok(registry.isRangeArg("AGGREGATE", 2), "Expected AGGREGATE ref1 to be a range");
  assert.ok(registry.isRangeArg("AGGREGATE", 3), "Expected AGGREGATE ref2 to be a range (varargs)");

  // CEILING.MATH/FLOOR.MATH: scalar-but-often-cell-referenced args should allow range/schema completion.
  assert.ok(registry.isRangeArg("CEILING.MATH", 0), "Expected CEILING.MATH number to be a range");
  assert.equal(registry.getFunction("CEILING.MATH")?.args?.[2]?.name, "mode", "Expected CEILING.MATH arg3 to be mode");
  assert.ok(registry.getFunction("CEILING.MATH")?.args?.[2]?.optional, "Expected CEILING.MATH mode to be optional");
  assert.ok(registry.isRangeArg("FLOOR.MATH", 0), "Expected FLOOR.MATH number to be a range");
  assert.ok(registry.isRangeArg("CEILING", 0), "Expected CEILING number to be a range");
  assert.equal(registry.getFunction("CEILING")?.args?.[1]?.name, "significance", "Expected CEILING arg2 to be significance");
  assert.ok(registry.isRangeArg("FLOOR", 0), "Expected FLOOR number to be a range");
  assert.ok(registry.isRangeArg("CEILING.PRECISE", 0), "Expected CEILING.PRECISE number to be a range");
  assert.ok(registry.getFunction("CEILING.PRECISE")?.args?.[1]?.optional, "Expected CEILING.PRECISE significance to be optional");
  assert.ok(registry.isRangeArg("FLOOR.PRECISE", 0), "Expected FLOOR.PRECISE number to be a range");
  assert.ok(registry.isRangeArg("ISO.CEILING", 0), "Expected ISO.CEILING number to be a range");
  assert.ok(registry.isRangeArg("MROUND", 0), "Expected MROUND number to be a range");
  assert.ok(registry.isRangeArg("EVEN", 0), "Expected EVEN number to be a range");
  assert.ok(registry.isRangeArg("ODD", 0), "Expected ODD number to be a range");

  // FORECAST.LINEAR(x, known_y's, known_x's)
  assert.equal(
    registry.isRangeArg("FORECAST.LINEAR", 0),
    false,
    "Expected FORECAST.LINEAR x not to be a range"
  );
  assert.equal(registry.getArgType("FORECAST.LINEAR", 0), "value", "Expected FORECAST.LINEAR x to be value-like");
  assert.ok(registry.isRangeArg("FORECAST.LINEAR", 1), "Expected FORECAST.LINEAR known_ys to be a range");
  assert.ok(registry.isRangeArg("FORECAST.LINEAR", 2), "Expected FORECAST.LINEAR known_xs to be a range");
  assert.equal(registry.getArgType("FORECAST", 0), "value", "Expected FORECAST x to be value-like");

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
  assert.equal(registry.getArgType("Z.TEST", 1), "value", "Expected Z.TEST x to be value-like");
  assert.equal(registry.getArgType("Z.TEST", 2), "value", "Expected Z.TEST sigma to be value-like");
  assert.ok(registry.getFunction("Z.TEST")?.args?.[2]?.optional, "Expected Z.TEST sigma to be optional");
  assert.equal(registry.getArgType("ZTEST", 1), "value", "Expected ZTEST x to be value-like");
  assert.equal(registry.getArgType("ZTEST", 2), "value", "Expected ZTEST sigma to be value-like");
  assert.ok(registry.getFunction("ZTEST")?.args?.[2]?.optional, "Expected ZTEST sigma to be optional");

  // Additional common stats functions
  assert.equal(registry.getArgType("LARGE", 1), "number", "Expected LARGE k to be numeric");
  assert.equal(registry.getArgType("SMALL", 1), "number", "Expected SMALL k to be numeric");
  assert.equal(registry.getArgType("PERCENTILE.INC", 1), "number", "Expected PERCENTILE.INC k to be numeric");
  assert.ok(registry.isRangeArg("PERCENTILE.EXC", 0), "Expected PERCENTILE.EXC array to be a range");
  assert.equal(registry.getArgType("PERCENTILE.EXC", 1), "number", "Expected PERCENTILE.EXC k to be numeric");
  assert.equal(registry.getArgType("PERCENTILE", 1), "number", "Expected PERCENTILE k to be numeric");
  assert.ok(registry.isRangeArg("QUARTILE.EXC", 0), "Expected QUARTILE.EXC array to be a range");
  assert.equal(registry.getArgType("RANK.EQ", 0), "value", "Expected RANK.EQ number to be value-like");
  assert.ok(registry.isRangeArg("RANK.AVG", 1), "Expected RANK.AVG ref to be a range");
  assert.equal(registry.getArgType("RANK.AVG", 0), "value", "Expected RANK.AVG number to be value-like");
  assert.ok(registry.isRangeArg("MODE.SNGL", 0), "Expected MODE.SNGL arg1 to be a range");
  assert.ok(registry.isRangeArg("TRIMMEAN", 0), "Expected TRIMMEAN array to be a range");
  assert.equal(registry.getArgType("TRIMMEAN", 1), "value", "Expected TRIMMEAN percent to be value-like");
  assert.equal(registry.getArgType("PERCENTRANK", 1), "value", "Expected PERCENTRANK x to be value-like");
  assert.equal(registry.getArgType("PERCENTRANK.INC", 1), "value", "Expected PERCENTRANK.INC x to be value-like");
  assert.equal(registry.getArgType("PERCENTRANK.EXC", 1), "value", "Expected PERCENTRANK.EXC x to be value-like");

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
  assert.equal(registry.getFunction("SEQUENCE")?.args?.[0]?.name, "rows", "Expected SEQUENCE arg1 to be rows");
  assert.ok(registry.getFunction("SEQUENCE")?.args?.[1]?.optional, "Expected SEQUENCE columns to be optional");
  assert.ok(registry.getFunction("RANDARRAY")?.args?.[0]?.optional, "Expected RANDARRAY rows to be optional");
  assert.equal(registry.getArgType("RANDARRAY", 4), "boolean", "Expected RANDARRAY whole_number to be boolean");
  assert.equal(registry.getFunction("RANDARRAY")?.args?.[4]?.name, "whole_number", "Expected RANDARRAY arg5 to be whole_number");
  assert.equal(registry.getArgType("SEQUENCE", 2), "value", "Expected SEQUENCE start to be value-like");
  assert.equal(registry.getArgType("SEQUENCE", 3), "value", "Expected SEQUENCE step to be value-like");
  assert.equal(registry.getArgType("RANDARRAY", 2), "value", "Expected RANDARRAY min to be value-like");
  assert.equal(registry.getArgType("RANDARRAY", 3), "value", "Expected RANDARRAY max to be value-like");

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
  assert.ok(registry.getFunction("AND")?.args?.[0]?.repeating, "Expected AND to accept repeating logical args");
  assert.ok(registry.getFunction("OR")?.args?.[0]?.repeating, "Expected OR to accept repeating logical args");
  assert.ok(registry.getFunction("XOR")?.args?.[0]?.repeating, "Expected XOR to accept repeating logical args");
  assert.equal(registry.getFunction("NOT")?.args?.[0]?.name, "logical", "Expected NOT arg1 name to be logical");
  assert.equal(registry.getFunction("CHOOSE")?.args?.[0]?.name, "index_num", "Expected CHOOSE arg1 name to be index_num");
  assert.ok(registry.getFunction("CHOOSE")?.args?.[1]?.repeating, "Expected CHOOSE value args to repeat");
  assert.equal(registry.getFunction("ERROR.TYPE")?.args?.[0]?.name, "error_val", "Expected ERROR.TYPE arg1 name to be error_val");

  // Legacy descriptive stats
  assert.ok(registry.isRangeArg("PERCENTILE", 0), "Expected PERCENTILE array to be a range");
  assert.ok(registry.isRangeArg("QUARTILE", 0), "Expected QUARTILE array to be a range");
  assert.ok(registry.isRangeArg("RANK", 1), "Expected RANK ref to be a range");
  assert.equal(registry.getArgType("RANK", 0), "value", "Expected RANK number to be value-like");
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
  assert.ok(registry.isRangeArg("INDEX", 0), "Expected INDEX array to be a range");
  assert.equal(registry.getArgType("INDEX", 1), "value", "Expected INDEX row_num to be value-like");
  assert.equal(registry.getArgType("INDEX", 2), "value", "Expected INDEX column_num to be value-like");
  assert.equal(registry.getArgType("INDEX", 3), "number", "Expected INDEX area_num to be numeric");

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
  assert.equal(registry.getArgType("FVSCHEDULE", 0), "value", "Expected FVSCHEDULE principal to be value-like");
  assert.ok(registry.isRangeArg("FVSCHEDULE", 1), "Expected FVSCHEDULE schedule to be a range");
  assert.ok(registry.isRangeArg("MIRR", 0), "Expected MIRR values to be a range");
  assert.equal(registry.getArgType("MIRR", 1), "value", "Expected MIRR finance_rate to be value-like");
  assert.equal(registry.getArgType("MIRR", 2), "value", "Expected MIRR reinvest_rate to be value-like");
  assert.ok(registry.isRangeArg("PROB", 0), "Expected PROB x_range to be a range");
  assert.ok(registry.isRangeArg("PROB", 1), "Expected PROB prob_range to be a range");
  assert.equal(registry.getArgType("PROB", 2), "value", "Expected PROB lower_limit to be value-like");
  assert.equal(registry.getArgType("PROB", 3), "value", "Expected PROB upper_limit to be value-like");
  assert.ok(registry.getFunction("PROB")?.args?.[3]?.optional, "Expected PROB upper_limit to be optional");
  assert.ok(registry.isRangeArg("SERIESSUM", 3), "Expected SERIESSUM coefficients to be a range");
  assert.equal(registry.getArgType("SERIESSUM", 0), "value", "Expected SERIESSUM x to be value-like");
  assert.equal(registry.getArgType("SERIESSUM", 1), "value", "Expected SERIESSUM n to be value-like");
  assert.equal(registry.getArgType("SERIESSUM", 2), "value", "Expected SERIESSUM m to be value-like");

  // Common distribution functions: keep cumulative flags boolean while treating scalar inputs as value-like.
  assert.equal(registry.getArgType("NORM.DIST", 0), "value", "Expected NORM.DIST x to be value-like");
  assert.equal(registry.getArgType("NORM.DIST", 3), "boolean", "Expected NORM.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("NORM.S.DIST", 0), "value", "Expected NORM.S.DIST z to be value-like");
  assert.equal(registry.getArgType("NORM.S.DIST", 1), "boolean", "Expected NORM.S.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("BINOM.DIST", 0), "value", "Expected BINOM.DIST number_s to be value-like");
  assert.equal(registry.getArgType("BINOM.DIST", 3), "boolean", "Expected BINOM.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("BINOMDIST", 0), "value", "Expected BINOMDIST number_s to be value-like");
  assert.equal(registry.getArgType("BINOMDIST", 3), "boolean", "Expected BINOMDIST cumulative to be boolean");
  assert.equal(registry.getArgType("HYPGEOM.DIST", 0), "value", "Expected HYPGEOM.DIST sample_s to be value-like");
  assert.equal(registry.getArgType("HYPGEOM.DIST", 4), "boolean", "Expected HYPGEOM.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("HYPGEOMDIST", 0), "value", "Expected HYPGEOMDIST sample_s to be value-like");
  assert.equal(registry.getArgType("NEGBINOM.DIST", 0), "value", "Expected NEGBINOM.DIST number_f to be value-like");
  assert.equal(registry.getArgType("NEGBINOM.DIST", 3), "boolean", "Expected NEGBINOM.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("NEGBINOMDIST", 0), "value", "Expected NEGBINOMDIST number_f to be value-like");
  assert.equal(registry.getArgType("BINOM.INV", 0), "value", "Expected BINOM.INV trials to be value-like");
  assert.equal(registry.getArgType("NORMDIST", 0), "value", "Expected NORMDIST x to be value-like");
  assert.equal(registry.getArgType("NORMDIST", 3), "boolean", "Expected NORMDIST cumulative to be boolean");
  assert.equal(registry.getArgType("NORMSDIST", 0), "value", "Expected NORMSDIST z to be value-like");
  assert.equal(registry.getArgType("NORMINV", 0), "value", "Expected NORMINV probability to be value-like");
  assert.equal(registry.getArgType("NORMSINV", 0), "value", "Expected NORMSINV probability to be value-like");
  assert.equal(registry.getArgType("LOGNORM.DIST", 0), "value", "Expected LOGNORM.DIST x to be value-like");
  assert.equal(registry.getArgType("LOGNORM.DIST", 3), "boolean", "Expected LOGNORM.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("EXPON.DIST", 0), "value", "Expected EXPON.DIST x to be value-like");
  assert.equal(registry.getArgType("EXPON.DIST", 2), "boolean", "Expected EXPON.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("GAMMA.DIST", 0), "value", "Expected GAMMA.DIST x to be value-like");
  assert.equal(registry.getArgType("GAMMA.DIST", 3), "boolean", "Expected GAMMA.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("BETA.DIST", 0), "value", "Expected BETA.DIST x to be value-like");
  assert.equal(registry.getArgType("BETA.DIST", 3), "boolean", "Expected BETA.DIST cumulative to be boolean");
  assert.ok(registry.getFunction("BETA.DIST")?.args?.[4]?.optional, "Expected BETA.DIST A to be optional");
  assert.equal(registry.getArgType("CHISQ.DIST", 0), "value", "Expected CHISQ.DIST x to be value-like");
  assert.equal(registry.getArgType("CHISQ.DIST", 2), "boolean", "Expected CHISQ.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("F.DIST", 0), "value", "Expected F.DIST x to be value-like");
  assert.equal(registry.getArgType("F.DIST", 3), "boolean", "Expected F.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("T.DIST", 0), "value", "Expected T.DIST x to be value-like");
  assert.equal(registry.getArgType("T.DIST", 2), "boolean", "Expected T.DIST cumulative to be boolean");
  assert.equal(registry.getArgType("TDIST", 0), "value", "Expected TDIST x to be value-like");
  assert.equal(registry.getArgType("TDIST", 2), "number", "Expected TDIST tails to be numeric");
  assert.equal(registry.getArgType("POISSON", 0), "value", "Expected POISSON x to be value-like");
  assert.equal(registry.getArgType("POISSON", 2), "boolean", "Expected POISSON cumulative to be boolean");
  assert.equal(registry.getArgType("WEIBULL", 0), "value", "Expected WEIBULL x to be value-like");
  assert.equal(registry.getArgType("WEIBULL", 3), "boolean", "Expected WEIBULL cumulative to be boolean");
  assert.equal(registry.getArgType("FISHER", 0), "value", "Expected FISHER x to be value-like");
  assert.equal(registry.getArgType("FISHERINV", 0), "value", "Expected FISHERINV y to be value-like");
  assert.equal(registry.getArgType("PHI", 0), "value", "Expected PHI x to be value-like");

  // Time-series forecasting functions
  assert.ok(registry.isRangeArg("FORECAST.ETS", 1), "Expected FORECAST.ETS values to be a range");
  assert.ok(registry.isRangeArg("FORECAST.ETS", 2), "Expected FORECAST.ETS timeline to be a range");

  // Date functions with optional holiday ranges
  assert.equal(registry.isRangeArg("WORKDAY", 0), false, "Expected WORKDAY start_date not to be a range");
  assert.equal(registry.getArgType("WORKDAY", 0), "value", "Expected WORKDAY start_date to be value-like");
  assert.equal(registry.isRangeArg("WORKDAY", 1), false, "Expected WORKDAY days not to be a range");
  assert.equal(registry.getArgType("WORKDAY", 1), "value", "Expected WORKDAY days to be value-like");
  assert.ok(registry.isRangeArg("WORKDAY", 2), "Expected WORKDAY holidays to be a range");
  assert.equal(registry.getArgType("WORKDAY.INTL", 1), "value", "Expected WORKDAY.INTL days to be value-like");
  assert.ok(registry.isRangeArg("WORKDAY.INTL", 3), "Expected WORKDAY.INTL holidays to be a range");
  assert.ok(registry.isRangeArg("NETWORKDAYS", 2), "Expected NETWORKDAYS holidays to be a range");
  assert.ok(registry.isRangeArg("NETWORKDAYS.INTL", 3), "Expected NETWORKDAYS.INTL holidays to be a range");
  assert.equal(registry.getArgType("NETWORKDAYS", 0), "value", "Expected NETWORKDAYS start_date to be value-like");
  assert.equal(registry.getArgType("NETWORKDAYS", 1), "value", "Expected NETWORKDAYS end_date to be value-like");

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
  assert.equal(registry.getArgType("ADDRESS", 0), "value", "Expected ADDRESS row_num to be value-like");
  assert.equal(registry.getArgType("ADDRESS", 1), "value", "Expected ADDRESS column_num to be value-like");
  assert.equal(registry.getArgType("NUMBERVALUE", 1), "string", "Expected NUMBERVALUE decimal_separator to be a string");
  assert.equal(registry.getArgType("NUMBERVALUE", 2), "string", "Expected NUMBERVALUE group_separator to be a string");
  assert.ok(registry.isRangeArg("NUMBERVALUE", 0), "Expected NUMBERVALUE text to be a range");
  assert.equal(registry.getArgType("IMAGE", 0), "string", "Expected IMAGE source to be a string");
  assert.equal(registry.getArgType("IMAGE", 2), "number", "Expected IMAGE sizing to be numeric");
  assert.equal(registry.getArgType("IMAGE", 3), "value", "Expected IMAGE height to be value-like");
  assert.equal(registry.getArgType("IMAGE", 4), "value", "Expected IMAGE width to be value-like");
  assert.ok(registry.isRangeArg("TEXT", 0), "Expected TEXT value to be a range");

  // Date/time helpers with more descriptive arg naming
  assert.equal(registry.getFunction("DATE")?.args?.[0]?.name, "year", "Expected DATE arg1 to be year");
  assert.equal(registry.getArgType("DATE", 0), "value", "Expected DATE year to be value-like");
  assert.equal(registry.getFunction("EDATE")?.args?.[0]?.name, "start_date", "Expected EDATE arg1 to be start_date");
  assert.equal(registry.getArgType("EDATE", 1), "value", "Expected EDATE months to be value-like");
  assert.equal(registry.getArgType("EOMONTH", 1), "value", "Expected EOMONTH months to be value-like");
  assert.equal(registry.getArgType("YEAR", 0), "value", "Expected YEAR serial_number to be value-like");
  assert.equal(registry.getArgType("WEEKDAY", 0), "value", "Expected WEEKDAY serial_number to be value-like");
  assert.equal(registry.getArgType("ISOWEEKNUM", 0), "value", "Expected ISOWEEKNUM serial_number to be value-like");
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
  assert.ok(registry.isRangeArg("DATEVALUE", 0), "Expected DATEVALUE date_text to be a range");
  assert.ok(registry.isRangeArg("TIMEVALUE", 0), "Expected TIMEVALUE time_text to be a range");

  // Core time value of money functions (catalog arg_types are too coarse; curated names improve hinting).
  const pv = registry.getFunction("PV");
  assert.ok(pv, "Expected PV to have a curated signature");
  assert.equal(pv?.args?.[0]?.name, "rate", "Expected PV arg1 to be rate");
  assert.equal(pv?.args?.[4]?.name, "type", "Expected PV arg5 to be type");
  assert.ok(pv?.args?.[4]?.optional, "Expected PV type to be optional");
  assert.equal(registry.getArgType("PV", 0), "value", "Expected PV rate to be value-like");
  assert.equal(registry.getArgType("PV", 1), "value", "Expected PV nper to be value-like");
  assert.equal(registry.getArgType("PV", 2), "value", "Expected PV pmt to be value-like");
  assert.equal(registry.getArgType("PV", 3), "value", "Expected PV fv to be value-like");
  assert.equal(registry.getArgType("PV", 4), "number", "Expected PV type to be numeric");
  assert.equal(registry.getArgType("NPV", 0), "value", "Expected NPV rate to be value-like");
  assert.equal(registry.getArgType("FV", 0), "value", "Expected FV rate to be value-like");
  assert.equal(registry.getArgType("PMT", 0), "value", "Expected PMT rate to be value-like");
  assert.equal(registry.getArgType("RATE", 5), "value", "Expected RATE guess to be value-like");
  assert.equal(registry.getArgType("IRR", 1), "value", "Expected IRR guess to be value-like");
  assert.equal(registry.getArgType("XNPV", 0), "value", "Expected XNPV rate to be value-like");
  assert.equal(registry.getArgType("XIRR", 2), "value", "Expected XIRR guess to be value-like");
  assert.equal(registry.getArgType("CUMIPMT", 5), "number", "Expected CUMIPMT type to be a number");
  assert.equal(registry.getArgType("CUMIPMT", 0), "value", "Expected CUMIPMT rate to be value-like");
  assert.equal(registry.getArgType("CUMPRINC", 0), "value", "Expected CUMPRINC rate to be value-like");
  assert.equal(registry.getArgType("VDB", 0), "value", "Expected VDB cost to be value-like");
  assert.equal(registry.getArgType("VDB", 5), "value", "Expected VDB factor to be value-like");
  assert.equal(registry.getFunction("VDB")?.args?.[6]?.type, "boolean", "Expected VDB no_switch to be boolean");

  // Bond/treasury functions: ensure arg naming matches enum indices in TabCompletionEngine.
  const price = registry.getFunction("PRICE");
  assert.ok(price, "Expected PRICE to have a curated signature");
  assert.equal(price?.args?.[0]?.name, "settlement", "Expected PRICE arg1 to be settlement");
  assert.equal(price?.args?.[5]?.name, "frequency", "Expected PRICE arg6 to be frequency");
  assert.ok(price?.args?.[6]?.optional, "Expected PRICE basis to be optional");
  assert.equal(registry.getArgType("PRICE", 2), "value", "Expected PRICE rate to be value-like");
  assert.equal(registry.getArgType("PRICE", 5), "number", "Expected PRICE frequency to stay numeric (enum hints)");
  assert.equal(registry.getArgType("PRICE", 6), "number", "Expected PRICE basis to stay numeric (enum hints)");

  assert.equal(registry.getArgType("ACCRINT", 7), "boolean", "Expected ACCRINT calc_method to be boolean");
  assert.equal(registry.getArgType("ACCRINT", 3), "value", "Expected ACCRINT rate to be value-like");
  assert.equal(registry.getArgType("ACCRINT", 5), "number", "Expected ACCRINT frequency to stay numeric (enum hints)");
  assert.equal(registry.getArgType("ACCRINT", 6), "number", "Expected ACCRINT basis to stay numeric (enum hints)");
  assert.equal(registry.getFunction("COUPDAYBS")?.args?.[2]?.name, "frequency", "Expected COUPDAYBS arg3 to be frequency");
  assert.equal(registry.getFunction("TBILLYIELD")?.args?.[2]?.name, "pr", "Expected TBILLYIELD arg3 to be pr");
  assert.equal(registry.getArgType("TBILLYIELD", 2), "value", "Expected TBILLYIELD pr to be value-like");
  assert.equal(registry.getArgType("INTRATE", 2), "value", "Expected INTRATE investment to be value-like");
  assert.equal(registry.getArgType("INTRATE", 4), "number", "Expected INTRATE basis to stay numeric (enum hints)");
  assert.equal(registry.getArgType("AMORLINC", 0), "value", "Expected AMORLINC cost to be value-like");
  assert.equal(registry.getArgType("AMORLINC", 6), "number", "Expected AMORLINC basis to stay numeric (enum hints)");
  assert.equal(registry.getArgType("DDB", 0), "value", "Expected DDB cost to be value-like");

  // Odd-period bond functions should keep frequency/basis positions aligned.
  assert.equal(registry.getFunction("ODDLPRICE")?.args?.[6]?.name, "frequency", "Expected ODDLPRICE arg7 to be frequency");
  assert.ok(registry.getFunction("ODDLPRICE")?.args?.[7]?.optional, "Expected ODDLPRICE basis to be optional");
  assert.equal(registry.getFunction("ODDFPRICE")?.args?.[7]?.name, "frequency", "Expected ODDFPRICE arg8 to be frequency");
  assert.ok(registry.getFunction("ODDFPRICE")?.args?.[8]?.optional, "Expected ODDFPRICE basis to be optional");
});
