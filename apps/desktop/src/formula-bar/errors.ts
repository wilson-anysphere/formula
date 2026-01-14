type ErrorExplanation = {
  code: string;
  title: string;
  description: string;
  suggestions: string[];
};

const ERROR_EXPLANATIONS: Record<string, Omit<ErrorExplanation, "code">> = {
  "#GETTING_DATA": {
    title: "Loading",
    description: "This cell is waiting for an async result (for example, an AI function response).",
    suggestions: ["Wait a moment for the result to arrive.", "If it never resolves, check your AI settings or network connection."],
  },
  "#DLP!": {
    title: "Blocked by data loss prevention",
    description: "This AI function call was blocked by your organization's DLP policy.",
    suggestions: [
      "Remove or change references to restricted cells/ranges.",
      "If this should be allowed, ask an admin to adjust the document/org DLP policy.",
    ],
  },
  "#AI!": {
    title: "AI error",
    description: "The AI function failed to run (model unavailable, network error, or unexpected response).",
    suggestions: ["Check your AI provider/model settings.", "Try again in a moment."],
  },
  "#NULL!": {
    title: "Null intersection",
    description: "The formula tried to reference an intersection that doesn't exist.",
    suggestions: ["Check that the referenced ranges actually intersect.", "Verify the formula’s range operators and separators."],
  },
  "#CALC!": {
    title: "Calculation error",
    description: "The formula couldn’t be calculated (often due to an unsupported or invalid operation).",
    suggestions: ["Check inputs for invalid values.", "Simplify the formula to isolate the failing part."],
  },
  "#FIELD!": {
    title: "Invalid field",
    description: "The formula referenced a field that doesn’t exist (often in data types or external data).",
    suggestions: ["Verify the field name exists.", "Refresh or re-import the underlying data."],
  },
  "#CONNECT!": {
    title: "Connection error",
    description: "The formula depends on external data that couldn’t be reached.",
    suggestions: ["Check your network connection.", "Try refreshing the data source."],
  },
  "#BLOCKED!": {
    title: "Blocked",
    description: "The formula result was blocked (for example, by a permission or data restriction).",
    suggestions: ["Check document permissions and data restrictions.", "Try moving the formula or adjusting inputs."],
  },
  "#UNKNOWN!": {
    title: "Unknown error",
    description: "The formula returned an unknown error.",
    suggestions: ["Try recalculating or re-entering the formula.", "If it persists, report a bug with the workbook."],
  },
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
