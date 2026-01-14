export type DecimalStepDirection = "increase" | "decrease";

function parseDecimalPlaces(format: string): number {
  const dot = format.indexOf(".");
  if (dot === -1) return 0;
  let count = 0;
  for (let i = dot + 1; i < format.length; i += 1) {
    const ch = format[i];
    if (ch === "0" || ch === "#") count += 1;
    else break;
  }
  return count;
}

/**
 * Best-effort helper for Excel-like "Increase Decimal" / "Decrease Decimal" behavior.
 *
 * This function intentionally supports only a limited subset of number format codes.
 *
 * It returns:
 * - a new number format string to apply, or
 * - `null` when stepping is not supported for the given format (fail closed).
 */
export function stepDecimalPlacesInNumberFormat(format: string | null, direction: DecimalStepDirection): string | null {
  const raw = (format ?? "").trim();
  const section = (raw.split(";")[0] ?? "").trim();
  const lower = section.toLowerCase();
  const compact = lower.replace(/\s+/g, "");

  // Avoid trying to manipulate date/time format codes.
  if (compact.includes("m/d/yyyy") || compact.includes("yyyy-mm-dd")) return null;
  if (/^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact)) return null;

  // Avoid mutating explicit text number formats.
  if (compact === "@") return null;

  // Preserve scientific notation when possible (e.g. `0.00E+00`).
  const scientificMatch = /E([+-])([0]+)/i.exec(section);
  if (scientificMatch) {
    const base = section.slice(0, scientificMatch.index);
    const decimals = parseDecimalPlaces(base);
    const nextDecimals = direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
    if (nextDecimals === decimals) return null;

    const expSign = scientificMatch[1] ?? "+";
    const expDigits = scientificMatch[2]?.length ?? 0;
    if (expDigits <= 0) return null;

    const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
    return `0${fraction}E${expSign}${"0".repeat(expDigits)}`;
  }

  // Preserve classic fraction formats (e.g. `# ?/?`, `# ??/??`) by adjusting the number of
  // `?` placeholders (instead of converting to a decimal format).
  const trimmed = section.trim();
  if (/^#\s+\?+\/\?+$/.test(trimmed)) {
    const slash = trimmed.indexOf("/");
    if (slash === -1) return null;
    const denom = trimmed.slice(slash + 1).trim();
    const digits = denom.length;
    const nextDigits = direction === "increase" ? Math.min(10, digits + 1) : Math.max(1, digits - 1);
    if (nextDigits === digits) return null;
    const qs = "?".repeat(nextDigits);
    return `# ${qs}/${qs}`;
  }

  const currencyMatch = /[$€£¥]/.exec(section);
  const prefix = currencyMatch?.[0] ?? "";
  const suffix = section.includes("%") ? "%" : "";
  const useThousands = section.includes(",");
  const decimals = parseDecimalPlaces(section);

  const nextDecimals = direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
  if (nextDecimals === decimals) return null;

  const integer = useThousands ? "#,##0" : "0";
  const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
  return `${prefix}${integer}${fraction}${suffix}`;
}

