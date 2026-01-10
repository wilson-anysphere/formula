import type { CalcSettings } from "./types";

export function formatCalculationStatus(
  settings: CalcSettings,
  opts: { dirty: boolean; circularRefs: number }
): string {
  const modeLabel =
    settings.calculationMode === "manual"
      ? "Manual"
      : settings.calculationMode === "automaticNoTable"
        ? "Automatic (No Tables)"
        : "Automatic";

  const parts: string[] = [modeLabel];
  if (opts.dirty && settings.calculationMode === "manual") {
    parts.push("Needs Calculation");
  }
  if (opts.circularRefs > 0 && !settings.iterative.enabled) {
    parts.push(`Circular (${opts.circularRefs})`);
  }
  return parts.join(" • ");
}
