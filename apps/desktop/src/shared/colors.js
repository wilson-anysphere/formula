/**
 * Normalize an Excel/OOXML color value into a CSS color string suitable for canvas rendering.
 *
 * Supported inputs:
 * - `#AARRGGBB` and `AARRGGBB` (Excel/OOXML ARGB)
 * - `#RRGGBB` and `RRGGBB`
 * - `#RGB` and `RGB` (CSS shorthand hex; expanded to `#RRGGBB`)
 * - Formula-model color reference objects:
 *   - `{ argb: string }`
 *   - `{ indexed: number }` (Excel indexed palette, 0..=63; 64 is "auto")
 *   - `{ theme: number, tint?: number }` (Office 2013 default theme palette)
 *   - `{ auto: true }`
 * - Any other non-hex string (e.g. `"red"`, `"rgb(…)"`) is returned as-is.
 *
 * Alpha is rounded to 3 decimal places for deterministic outputs (e.g. `0x80 / 255` -> `0.502`).
 *
 * @param {unknown} input
 * @returns {string | undefined}
 */
export function normalizeExcelColorToCss(input) {
  if (typeof input === "string") {
    return normalizeExcelColorStringToCss(input);
  }

  // formula-model serializes Excel colors as either:
  // - "#AARRGGBB" strings (direct ARGB), OR
  // - objects like `{ theme, tint }`, `{ indexed }`, `{ auto: true }`
  // Additionally, we sometimes see `{ argb: "#AARRGGBB" }` from other XLSX tooling.
  if (!isPlainObject(input)) return undefined;

  // `{"argb":"#AARRGGBB"}`
  if (typeof input.argb === "string") {
    return normalizeExcelColorStringToCss(input.argb);
  }

  // `{ auto: true }` is context-dependent ("automatic") and cannot be resolved here.
  if (input.auto === true) return undefined;

  // DocumentController + desktop sheet metadata store represent tab colors as `{ rgb }`.
  if (typeof input.rgb === "string") {
    return normalizeExcelColorStringToCss(input.rgb);
  }

  // `{ indexed: number }`
  if (typeof input.indexed === "number" && Number.isFinite(input.indexed)) {
    const idx = input.indexed;
    // Excel (BIFF8) reserves index 64 as "automatic".
    if (idx === 64) return undefined;

    if (!Number.isInteger(idx) || idx < 0 || idx >= EXCEL_INDEXED_COLORS.length) return undefined;
    const argb = EXCEL_INDEXED_COLORS[idx];
    return normalizeExcelColorStringToCss(toHexArgb(argb));
  }

  // `{ theme: number, tint?: number }`
  if (typeof input.theme === "number" && Number.isFinite(input.theme)) {
    const themeIndex = input.theme;
    if (!Number.isInteger(themeIndex) || themeIndex < 0 || themeIndex >= OFFICE_2013_THEME_PALETTE.length) return undefined;

    let argb = OFFICE_2013_THEME_PALETTE[themeIndex];

    if (typeof input.tint === "number" && Number.isFinite(input.tint) && input.tint !== 0) {
      argb = applyTint(argb, input.tint);
    }

    return normalizeExcelColorStringToCss(toHexArgb(argb));
  }

  return undefined;
}

/**
 * @param {string} input
 * @returns {string | undefined}
 */
function normalizeExcelColorStringToCss(input) {
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
  if (raw.length === 3) {
    // CSS shorthand hex: #RGB / RGB
    const [r, g, b] = raw.toLowerCase().split("");
    return `#${r}${r}${g}${g}${b}${b}`;
  }

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

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown>}
 */
function isPlainObject(value) {
  if (value === null || typeof value !== "object") return false;
  if (Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/**
 * @param {number} argb
 * @returns {string}
 */
function toHexArgb(argb) {
  // Ensure we treat the number as an unsigned 32-bit integer even if it has the
  // high bit set (JS numbers are signed when coerced to int32).
  return (argb >>> 0).toString(16).padStart(8, "0");
}

/**
 * Port of `crates/formula-model/src/theme.rs` tint logic.
 *
 * @param {number} argb
 * @param {number} tintThousandths
 * @returns {number}
 */
function applyTint(argb, tintThousandths) {
  const tint = clamp(tintThousandths / 1000, -1, 1);
  if (tint === 0) return argb;

  const a = (argb >>> 24) & 0xff;
  const r = (argb >>> 16) & 0xff;
  const g = (argb >>> 8) & 0xff;
  const b = argb & 0xff;

  const outR = tintChannel(r, tint);
  const outG = tintChannel(g, tint);
  const outB = tintChannel(b, tint);

  return ((a << 24) | (outR << 16) | (outG << 8) | outB) >>> 0;
}

/**
 * @param {number} value
 * @param {number} tint
 * @returns {number}
 */
function tintChannel(value, tint) {
  const v = value;
  const out = tint < 0 ? v * (1 + tint) : v * (1 - tint) + 255 * tint;
  return clamp(Math.round(out), 0, 255);
}

/**
 * @param {number} v
 * @param {number} min
 * @param {number} max
 * @returns {number}
 */
function clamp(v, min, max) {
  return Math.max(min, Math.min(max, v));
}

// Excel's standard indexed color table for indices `0..=63`.
// Reference: `crates/formula-model/src/theme.rs` (`EXCEL_INDEXED_COLORS`).
const EXCEL_INDEXED_COLORS = [
  0xff000000, // 0
  0xffffffff, // 1
  0xffff0000, // 2
  0xff00ff00, // 3
  0xff0000ff, // 4
  0xffffff00, // 5
  0xffff00ff, // 6
  0xff00ffff, // 7
  0xff000000, // 8
  0xffffffff, // 9
  0xffff0000, // 10
  0xff00ff00, // 11
  0xff0000ff, // 12
  0xffffff00, // 13
  0xffff00ff, // 14
  0xff00ffff, // 15
  0xff800000, // 16
  0xff008000, // 17
  0xff000080, // 18
  0xff808000, // 19
  0xff800080, // 20
  0xff008080, // 21
  0xffc0c0c0, // 22
  0xff808080, // 23
  0xff9999ff, // 24
  0xff993366, // 25
  0xffffffcc, // 26
  0xffccffff, // 27
  0xff660066, // 28
  0xffff8080, // 29
  0xff0066cc, // 30
  0xffccccff, // 31
  0xff000080, // 32
  0xffff00ff, // 33
  0xffffff00, // 34
  0xff00ffff, // 35
  0xff800080, // 36
  0xff800000, // 37
  0xff008080, // 38
  0xff0000ff, // 39
  0xff00ccff, // 40
  0xffccffff, // 41
  0xffccffcc, // 42
  0xffffff99, // 43
  0xff99ccff, // 44
  0xffff99cc, // 45
  0xffcc99ff, // 46
  0xffffcc99, // 47
  0xff3366ff, // 48
  0xff33cccc, // 49
  0xff99cc00, // 50
  0xffffcc00, // 51
  0xffff9900, // 52
  0xffff6600, // 53
  0xff666699, // 54
  0xff969696, // 55
  0xff003366, // 56
  0xff339966, // 57
  0xff003300, // 58
  0xff333300, // 59
  0xff993300, // 60
  0xff993366, // 61
  0xff333399, // 62
  0xff333333, // 63
];

// Office 2013+ default theme palette (SpreadsheetML theme indices 0..=11).
// Reference: `crates/formula-model/src/theme.rs` (`ThemePalette::office_2013`).
const OFFICE_2013_THEME_PALETTE = [
  0xffffffff, // 0: lt1
  0xff000000, // 1: dk1
  0xffe7e6e6, // 2: lt2
  0xff44546a, // 3: dk2
  0xff5b9bd5, // 4: accent1
  0xffed7d31, // 5: accent2
  0xffa5a5a5, // 6: accent3
  0xffffc000, // 7: accent4
  0xff4472c4, // 8: accent5
  0xff70ad47, // 9: accent6
  0xff0563c1, // 10: hlink
  0xff954f72, // 11: folHlink
];
