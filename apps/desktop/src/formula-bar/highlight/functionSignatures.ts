export type FunctionParam = { name: string; optional?: boolean };

export type FunctionSignature = {
  name: string;
  params: FunctionParam[];
  summary: string;
};

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
  return FUNCTION_SIGNATURES[name.toUpperCase()] ?? null;
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
