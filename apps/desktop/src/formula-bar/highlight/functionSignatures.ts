import FUNCTION_CATALOG from "../../../../../shared/functionCatalog.mjs";

export type FunctionParam = { name: string; optional?: boolean };

export type FunctionSignature = {
  name: string;
  params: FunctionParam[];
  summary: string;
};

type CatalogFunction = {
  name: string;
  min_args: number;
  max_args: number;
  arg_types?: string[];
};

const CATALOG_BY_NAME = new Map<string, CatalogFunction>();
for (const fn of (FUNCTION_CATALOG as { functions?: CatalogFunction[] } | null)?.functions ?? []) {
  if (fn?.name) CATALOG_BY_NAME.set(fn.name.toUpperCase(), fn);
}

export const FUNCTION_SIGNATURES: Record<string, FunctionSignature> = {
  SUM: {
    name: "SUM",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Adds all the numbers in a range of cells.",
  },
  COUNT: {
    name: "COUNT",
    params: [{ name: "value1" }, { name: "value2", optional: true }],
    summary: "Counts the number of cells that contain numbers.",
  },
  COUNTA: {
    name: "COUNTA",
    params: [{ name: "value1" }, { name: "value2", optional: true }],
    summary: "Counts the number of non-empty cells.",
  },
  COUNTBLANK: {
    name: "COUNTBLANK",
    params: [{ name: "range" }],
    summary: "Counts the number of blank cells within a range.",
  },
  AVERAGE: {
    name: "AVERAGE",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the average (arithmetic mean) of its arguments.",
  },
  MAX: {
    name: "MAX",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the largest value in a set of values.",
  },
  MIN: {
    name: "MIN",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the smallest value in a set of values.",
  },
  IF: {
    name: "IF",
    params: [
      { name: "logical_test" },
      { name: "value_if_true" },
      { name: "value_if_false", optional: true },
    ],
    summary: "Checks whether a condition is met, and returns one value if TRUE and another value if FALSE.",
  },
  VLOOKUP: {
    name: "VLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "table_array" },
      { name: "col_index_num" },
      { name: "range_lookup", optional: true },
    ],
    summary: "Looks for a value in the leftmost column of a table, then returns a value in the same row from a specified column.",
  },
  HLOOKUP: {
    name: "HLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "table_array" },
      { name: "row_index_num" },
      { name: "range_lookup", optional: true },
    ],
    summary: "Looks for a value in the top row of a table, then returns a value in the same column from a specified row.",
  },
  XLOOKUP: {
    name: "XLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "lookup_array" },
      { name: "return_array" },
      { name: "if_not_found", optional: true },
      { name: "match_mode", optional: true },
      { name: "search_mode", optional: true },
    ],
    summary: "Looks up a value in a range or an array.",
  },
  INDEX: {
    name: "INDEX",
    params: [
      { name: "array" },
      { name: "row_num" },
      { name: "column_num", optional: true },
    ],
    summary: "Returns the value of an element in a table or an array.",
  },
  MATCH: {
    name: "MATCH",
    params: [
      { name: "lookup_value" },
      { name: "lookup_array" },
      { name: "match_type", optional: true },
    ],
    summary: "Looks up values in a reference or array.",
  },
  TODAY: {
    name: "TODAY",
    params: [],
    summary: "Returns the current date.",
  },
  NOW: {
    name: "NOW",
    params: [],
    summary: "Returns the current date and time.",
  },
  TRANSPOSE: {
    name: "TRANSPOSE",
    params: [{ name: "array" }],
    summary: "Returns the transpose of an array or range.",
  },
};

export function getFunctionSignature(name: string): FunctionSignature | null {
  const requested = name.toUpperCase();
  const lookup = requested.startsWith("_XLFN.") ? requested.slice("_XLFN.".length) : requested;

  const known = FUNCTION_SIGNATURES[lookup] ?? signatureFromCatalog(lookup);
  if (!known) return null;

  // Preserve any `_xlfn.` prefix in the displayed name so formula-bar hints
  // match pasted formulas from Excel files.
  return lookup === requested ? known : { ...known, name: requested };
}

export type SignaturePart = { text: string; kind: "name" | "param" | "paramActive" | "punct" };

export function signatureParts(sig: FunctionSignature, activeParamIndex: number | null): SignaturePart[] {
  const parts: SignaturePart[] = [{ text: `${sig.name}(`, kind: "name" }];
  sig.params.forEach((param, index) => {
    if (index > 0) parts.push({ text: ", ", kind: "punct" });
    const isActive = activeParamIndex !== null && activeParamIndex === index;
    parts.push({
      text: param.optional ? `[${param.name}]` : param.name,
      kind: isActive ? "paramActive" : "param",
    });
  });
  parts.push({ text: ")", kind: "punct" });
  return parts;
}

function signatureFromCatalog(name: string): FunctionSignature | null {
  const fn = CATALOG_BY_NAME.get(name);
  if (!fn) return null;

  return {
    name,
    params: buildParams(fn.min_args, fn.max_args, fn.arg_types),
    summary: "",
  };
}

function buildParams(minArgs: number, maxArgs: number, argTypes: string[] | undefined): FunctionParam[] {
  const MAX_PARAMS = 5;

  if (!Number.isFinite(minArgs) || !Number.isFinite(maxArgs) || minArgs < 0 || maxArgs < 0) {
    return [];
  }

  if (maxArgs <= MAX_PARAMS) {
    const out: FunctionParam[] = [];
    for (let i = 1; i <= maxArgs; i++) {
      out.push({ name: paramNameFromCatalogTypes(i, maxArgs, argTypes), optional: i > minArgs });
    }
    return out;
  }

  const requiredShown = Math.min(minArgs, MAX_PARAMS - 1);
  const out: FunctionParam[] = [];
  for (let i = 1; i <= requiredShown; i++) out.push({ name: paramNameFromCatalogTypes(i, maxArgs, argTypes) });

  if (minArgs > requiredShown) {
    out.push({ name: "…" });
    return out;
  }

  out.push({ name: "…", optional: true });
  return out;
}

function paramNameFromCatalogTypes(index1: number, maxArgs: number, argTypes: string[] | undefined): string {
  const index0 = index1 - 1;
  if (!Array.isArray(argTypes) || argTypes.length === 0) return `arg${index1}`;

  let valueType: string | undefined;
  if (argTypes.length === 1 && maxArgs > 1) {
    valueType = argTypes[0];
  } else {
    valueType = argTypes[index0] ?? argTypes[argTypes.length - 1];
  }

  switch (valueType) {
    case "number":
      return `number${index1}`;
    case "text":
      return `text${index1}`;
    case "bool":
      return `logical${index1}`;
    case "any":
      return `value${index1}`;
    default:
      return `arg${index1}`;
  }
}
