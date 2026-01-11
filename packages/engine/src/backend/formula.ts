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
  if (typeof value !== "string") return false;
  // We intentionally treat leading whitespace before '=' as formula input because
  // the engine protocol uses a leading '=' to disambiguate formulas from literal
  // text. However, a bare "=" / "=   " is treated as literal text (not a
  // formula), matching the WASM engine's detection rules.
  const trimmed = value.trimStart();
  if (!trimmed.startsWith("=")) return false;
  return trimmed.slice(1).trim() !== "";
}
