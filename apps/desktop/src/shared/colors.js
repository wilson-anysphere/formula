/**
 * Normalize an Excel/OOXML color value into a CSS color string suitable for canvas rendering.
 *
 * Supported inputs:
 * - `#AARRGGBB` and `AARRGGBB` (Excel/OOXML ARGB)
 * - `#RRGGBB` and `RRGGBB`
 * - Any other non-hex string (e.g. `"red"`, `"rgb(…)"`) is returned as-is.
 *
 * Alpha is rounded to 3 decimal places for deterministic outputs (e.g. `0x80 / 255` -> `0.502`).
 *
 * @param {unknown} input
 * @returns {string | undefined}
 */
export function normalizeExcelColorToCss(input) {
  if (typeof input !== "string") return undefined;

  const trimmed = input.trim();
  if (!trimmed) return undefined;

  const hasHash = trimmed.startsWith("#");
  const raw = hasHash ? trimmed.slice(1) : trimmed;

  // If the value doesn't look like a hex token, assume it's a CSS color string
  // (named colors, rgb()/rgba(), var(--foo), etc) and return it as-is.
  if (!hasHash && !/^[0-9a-fA-F]+$/.test(raw)) return trimmed;

  // At this point the caller either used a `#` prefix or the token is hex-like.
  // Treat invalid hex chars as invalid input rather than passing through.
  if (!/^[0-9a-fA-F]+$/.test(raw)) return undefined;

  // #RRGGBB / RRGGBB
  if (raw.length === 6) {
    return `#${raw.toLowerCase()}`;
  }

  // Excel/OOXML ARGB: #AARRGGBB / AARRGGBB
  if (raw.length === 8) {
    const a = Number.parseInt(raw.slice(0, 2), 16);
    const r = Number.parseInt(raw.slice(2, 4), 16);
    const g = Number.parseInt(raw.slice(4, 6), 16);
    const b = Number.parseInt(raw.slice(6, 8), 16);

    if (![a, r, g, b].every((n) => Number.isFinite(n))) return undefined;

    if (a >= 255) {
      return `#${raw.slice(2).toLowerCase()}`;
    }

    const alpha = Math.max(0, Math.min(1, a / 255));
    const rounded = Math.round(alpha * 1000) / 1000;
    const alphaStr = rounded.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
    return `rgba(${r},${g},${b},${alphaStr})`;
  }

  return undefined;
}
