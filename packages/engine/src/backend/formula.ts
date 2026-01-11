export function normalizeFormulaText(formula: string): string {
  const trimmed = formula.trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();

  if (stripped === "") return "";
  return `=${stripped}`;
}

export function normalizeFormulaTextOpt(formula: string): string | null {
  const normalized = normalizeFormulaText(formula);
  return normalized === "" ? null : normalized;
}

export function isFormulaInput(value: unknown): value is string {
  // We intentionally treat leading whitespace before '=' as "formula input"
  // because the engine-side protocol uses a leading '=' to disambiguate formula
  // strings from literal text scalars.
  return typeof value === "string" && value.trimStart().startsWith("=");
}
