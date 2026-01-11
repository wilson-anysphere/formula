import type { CellScalar } from "@formula/engine";
import { normalizeFormulaTextOpt } from "@formula/engine";

export function scalarToDisplayString(value: CellScalar): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function parseCellScalarInput(raw: string): CellScalar {
  if (raw.trimStart().startsWith("=")) {
    return normalizeFormulaTextOpt(raw);
  }

  const trimmed = raw.trim();
  if (trimmed === "") return null;

  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  if (/^null$/i.test(trimmed)) return null;

  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }

  return raw;
}

export function isFormulaInput(value: CellScalar): value is string {
  return typeof value === "string" && value.trimStart().startsWith("=");
}
