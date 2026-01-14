/**
 * Normalize an Excel/OOXML color value into a CSS color string suitable for canvas rendering.
 *
 * Supported inputs:
 * - `#AARRGGBB` and `AARRGGBB` (Excel/OOXML ARGB)
 * - `#RRGGBB` and `RRGGBB`
 * - `#RGB` and `RGB` (CSS shorthand hex; expanded to `#RRGGBB`)
 * - XLSX/OOXML color reference objects (formula-model + common tooling variants):
 *   - `{ argb: string, tint?: number }`
 *   - `{ rgb: string, tint?: number }`
 *   - `{ indexed: number, tint?: number }` (Excel indexed palette, 0..=63; 64 is "auto")
 *   - `{ theme: number, tint?: number }` (Office 2013 default theme palette)
 *   - `{ auto: true }`
 * - Any other non-hex string (e.g. `"red"`, `"rgb(…)"`) is returned as-is.
 *
 * Tint values may be expressed either as:
 * - thousandths in `[-1000, 1000]` (formula-model, e.g. `-500`), or
 * - fractions in `[-1, 1]` (OOXML-style, e.g. `-0.5`).
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
    const base = normalizeExcelColorStringToCss(input.argb);
    if (typeof input.tint === "number" && Number.isFinite(input.tint) && input.tint !== 0) {
      const parsed = parseExcelColorStringToArgb(input.argb);
      if (parsed != null) {
        const tinted = applyTint(parsed, input.tint);
        return normalizeExcelColorStringToCss(toHexArgb(tinted));
      }
    }
    return base;
  }

  // `{ auto: true }` is context-dependent ("automatic") and cannot be resolved here.
  if (input.auto === true) return undefined;

  // DocumentController + desktop sheet metadata store represent tab colors as `{ rgb }`.
  if (typeof input.rgb === "string") {
    const base = normalizeExcelColorStringToCss(input.rgb);
    if (typeof input.tint === "number" && Number.isFinite(input.tint) && input.tint !== 0) {
      const parsed = parseExcelColorStringToArgb(input.rgb);
      if (parsed != null) {
        const tinted = applyTint(parsed, input.tint);
        return normalizeExcelColorStringToCss(toHexArgb(tinted));
      }
    }
    return base;
  }

  // `{ indexed: number }`
  if (typeof input.indexed === "number" && Number.isFinite(input.indexed)) {
    const idx = input.indexed;
    // Excel (BIFF8) reserves index 64 as "automatic".
    if (idx === 64) return undefined;

    if (!Number.isInteger(idx) || idx < 0 || idx >= EXCEL_INDEXED_COLORS.length) return undefined;
    let argb = EXCEL_INDEXED_COLORS[idx];
    if (typeof input.tint === "number" && Number.isFinite(input.tint) && input.tint !== 0) {
      argb = applyTint(argb, input.tint);
    }
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

// Normalize/parse of Excel color strings (ARGB/hex) is on hot render paths (grid, rich text).
// Cache the output for repeated calls with the same token to avoid repeated parsing/formatting.
//
// Cap the cache so we don't grow unbounded when documents contain many unique colors.
const COLOR_STRING_CACHE_MAX = 4096;
/** @type {Map<string, string | null>} */
const COLOR_STRING_CACHE = new Map();

/**
 * @param {string} input
 * @returns {string | undefined}
 */
function normalizeExcelColorStringToCss(input) {
  const trimmed = input.trim();
  if (!trimmed) return undefined;

  const cached = COLOR_STRING_CACHE.get(trimmed);
  if (cached !== undefined) return cached ?? undefined;

  const hasHash = trimmed.startsWith("#");
  const raw = hasHash ? trimmed.slice(1) : trimmed;

  // If the value doesn't look like a hex token, assume it's a CSS color string
  // (named colors, rgb()/rgba(), var(--foo), etc) and return it as-is.
  if (!hasHash && !/^[0-9a-fA-F]+$/.test(raw)) {
    COLOR_STRING_CACHE.set(trimmed, trimmed);
    if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
    return trimmed;
  }

  // At this point the caller either used a `#` prefix or the token is hex-like.
  // Treat invalid hex chars as invalid input rather than passing through.
  if (!/^[0-9a-fA-F]+$/.test(raw)) {
    COLOR_STRING_CACHE.set(trimmed, null);
    if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
    return undefined;
  }

  // #RRGGBB / RRGGBB
  if (raw.length === 3) {
    // CSS shorthand hex: #RGB / RGB
    const [r, g, b] = raw.toLowerCase().split("");
    const out = `#${r}${r}${g}${g}${b}${b}`;
    COLOR_STRING_CACHE.set(trimmed, out);
    if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
    return out;
  }

  if (raw.length === 6) {
    const out = `#${raw.toLowerCase()}`;
    COLOR_STRING_CACHE.set(trimmed, out);
    if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
    return out;
  }

  // Excel/OOXML ARGB: #AARRGGBB / AARRGGBB
  if (raw.length === 8) {
    const a = Number.parseInt(raw.slice(0, 2), 16);
    const r = Number.parseInt(raw.slice(2, 4), 16);
    const g = Number.parseInt(raw.slice(4, 6), 16);
    const b = Number.parseInt(raw.slice(6, 8), 16);

    if (![a, r, g, b].every((n) => Number.isFinite(n))) return undefined;

    if (a >= 255) {
      const out = `#${raw.slice(2).toLowerCase()}`;
      COLOR_STRING_CACHE.set(trimmed, out);
      if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
      return out;
    }

    const alpha = Math.max(0, Math.min(1, a / 255));
    const rounded = Math.round(alpha * 1000) / 1000;
    const alphaStr = rounded.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
    const out = `rgba(${r},${g},${b},${alphaStr})`;
    COLOR_STRING_CACHE.set(trimmed, out);
    if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
    return out;
  }

  COLOR_STRING_CACHE.set(trimmed, null);
  if (COLOR_STRING_CACHE.size > COLOR_STRING_CACHE_MAX) COLOR_STRING_CACHE.clear();
  return undefined;
}

/**
 * Parse an Excel/OOXML-style hex color string into an ARGB integer.
 *
 * Supports:
 * - `#AARRGGBB` / `AARRGGBB`
 * - `#RRGGBB` / `RRGGBB`
 * - `#RGB` / `RGB` (expanded to `RRGGBB`)
 *
 * Returns `null` when the input is not a hex-like token.
 *
 * @param {string} input
 * @returns {number | null}
 */
function parseExcelColorStringToArgb(input) {
  const trimmed = input.trim();
  if (!trimmed) return null;

  const hasHash = trimmed.startsWith("#");
  const raw = hasHash ? trimmed.slice(1) : trimmed;

  // If the value doesn't look like a hex token, treat it as a CSS string that we can't parse.
  if (!hasHash && !/^[0-9a-fA-F]+$/.test(raw)) return null;
  if (!/^[0-9a-fA-F]+$/.test(raw)) return null;

  if (raw.length === 3) {
    const [r, g, b] = raw.toLowerCase().split("");
    const expanded = `${r}${r}${g}${g}${b}${b}`;
    const rgb = Number.parseInt(expanded, 16);
    if (!Number.isFinite(rgb)) return null;
    return (0xff << 24) | rgb;
  }

  if (raw.length === 6) {
    const rgb = Number.parseInt(raw, 16);
    if (!Number.isFinite(rgb)) return null;
    return (0xff << 24) | rgb;
  }

  if (raw.length === 8) {
    const argb = Number.parseInt(raw, 16);
    if (!Number.isFinite(argb)) return null;
    // Ensure unsigned.
    return argb >>> 0;
  }

  return null;
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
 * `tint` values in the wild come in two scales:
 * - Formula-model uses thousandths (e.g. -500 == -0.5)
 * - OOXML (and some tooling) use fractions in [-1, 1] (e.g. -0.5)
 *
 * Accept either by auto-detecting the scale based on magnitude.
 *
 * @param {number} argb
 * @param {number} tintValue
 * @returns {number}
 */
function applyTint(argb, tintValue) {
  let tint = Number(tintValue);
  if (!Number.isFinite(tint)) return argb;
  if (Math.abs(tint) > 1) {
    tint /= 1000;
  }
  tint = clamp(tint, -1, 1);
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
