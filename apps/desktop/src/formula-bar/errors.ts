export type ErrorExplanation = {
  code: string;
  title: string;
  description: string;
  suggestions: string[];
};

const ERROR_EXPLANATIONS: Record<string, Omit<ErrorExplanation, "code">> = {
  "#DIV/0!": {
    title: "Division by zero",
    description: "The formula tried to divide by zero (or an empty cell).",
    suggestions: [
      "Check the divisor cell for a 0 or blank value.",
      "Wrap the division in IFERROR to provide a fallback value.",
    ],
  },
  "#NAME?": {
    title: "Unknown name",
    description: "The formula contains a function name or named range that isn’t recognized.",
    suggestions: ["Check the spelling of function names.", "Verify that referenced named ranges exist."],
  },
  "#REF!": {
    title: "Invalid reference",
    description: "The formula refers to a cell or range that no longer exists.",
    suggestions: ["Check for deleted rows/columns in referenced ranges.", "Update the formula to point to valid cells."],
  },
  "#VALUE!": {
    title: "Wrong type of value",
    description: "The formula used a value of the wrong type (e.g. text where a number was expected).",
    suggestions: ["Check referenced cells for unexpected text values.", "Use VALUE or other coercion helpers if needed."],
  },
  "#N/A": {
    title: "Value not available",
    description: "A lookup didn’t find a matching value (or data is missing).",
    suggestions: ["Verify the lookup value exists in the lookup range.", "Consider IFNA/IFERROR to handle missing values."],
  },
  "#NUM!": {
    title: "Invalid number",
    description: "The formula produced an invalid numeric result (too large/small or not representable).",
    suggestions: ["Check for invalid inputs (like negative numbers where not allowed).", "Simplify the calculation to avoid overflow."],
  },
  "#SPILL!": {
    title: "Spill range blocked",
    description: "A dynamic array formula can’t spill because cells in the spill area are not empty.",
    suggestions: ["Clear the cells where the formula needs to spill.", "Move the formula to an empty area."],
  },
};

export function explainFormulaError(value: unknown): ErrorExplanation | null {
  if (typeof value !== "string") return null;
  const explanation = ERROR_EXPLANATIONS[value];
  if (!explanation) return null;
  return { code: value, ...explanation };
}

