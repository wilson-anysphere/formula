export type FunctionParam = { name: string; optional?: boolean };

export type FunctionSignature = {
  name: string;
  params: FunctionParam[];
  summary: string;
};

import FUNCTION_CATALOG from "../../../../../shared/functionCatalog.json" with { type: "json" };

type CatalogFunction = {
  name: string;
  min_args: number;
  max_args: number;
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
  AVERAGE: {
    name: "AVERAGE",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the average (arithmetic mean) of its arguments.",
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
    params: buildParams(fn.min_args, fn.max_args),
    summary: "",
  };
}

function buildParams(minArgs: number, maxArgs: number): FunctionParam[] {
  const MAX_PARAMS = 5;

  if (!Number.isFinite(minArgs) || !Number.isFinite(maxArgs) || minArgs < 0 || maxArgs < 0) {
    return [];
  }

  if (maxArgs <= MAX_PARAMS) {
    const out: FunctionParam[] = [];
    for (let i = 1; i <= maxArgs; i++) {
      out.push({ name: `arg${i}`, optional: i > minArgs });
    }
    return out;
  }

  const requiredShown = Math.min(minArgs, MAX_PARAMS - 1);
  const out: FunctionParam[] = [];
  for (let i = 1; i <= requiredShown; i++) out.push({ name: `arg${i}` });

  if (minArgs > requiredShown) {
    out.push({ name: "…" });
    return out;
  }

  out.push({ name: "…", optional: true });
  return out;
}
