import type { TraceNode } from "./types.ts";

export type ErrorExplanation = {
  title: string;
  problem: string;
  suggestions: string[];
};

function asError(value: unknown): string | null {
  if (typeof value === "string" && value.startsWith("#")) return value;
  if (value && typeof value === "object" && "error" in value) {
    const e = (value as { error?: unknown }).error;
    if (typeof e === "string") return e;
  }
  return null;
}

export function explainError(formula: string, trace: TraceNode): ErrorExplanation | null {
  const rootError = asError(trace.value);
  if (!rootError) return null;

  if (rootError === "#N/A") {
    const vlookup = findVlookupError(trace);
    if (vlookup) {
      const lookupValue = vlookup.children?.[0]?.value;
      const rangeRef = vlookup.children?.[1]?.reference;
      const range = rangeRef && rangeRef.type === "range" ? rangeRef.range : "the lookup table";
      return {
        title: `⚠️ ${rootError} Error`,
        problem: `The lookup value ${JSON.stringify(lookupValue)} was not found in the first column of ${range}.`,
        suggestions: [
          "Check if the lookup value exists in the lookup range.",
          "Verify the lookup range is correct and includes the expected rows.",
          "Consider wrapping with IFERROR to handle missing values.",
        ],
      };
    }
  }

  return {
    title: `⚠️ ${rootError} Error`,
    problem: `Formula ${formula} evaluated to ${rootError}.`,
    suggestions: ["Inspect intermediate steps in the formula debugger to find where the error originates."],
  };
}

function findVlookupError(node: TraceNode): TraceNode | null {
  const err = asError(node.value);
  if (
    err === "#N/A" &&
    node.kind.type === "function_call" &&
    node.kind.name.toUpperCase() === "VLOOKUP"
  ) {
    return node;
  }
  for (const child of node.children ?? []) {
    const found = findVlookupError(child);
    if (found) return found;
  }
  return null;
}

