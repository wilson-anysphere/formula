export function normalizeFormulaText(formula: string): string {
  const trimmed = formula.trimStart();
  if (trimmed.startsWith("=")) return trimmed;
  return `=${trimmed}`;
}

export function isFormulaInput(value: unknown): value is string {
  return typeof value === "string" && value.trimStart().startsWith("=");
}

