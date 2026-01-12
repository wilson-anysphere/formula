/**
 * RTF clipboard serialization.
 *
 * We emit an RTF table so rich clipboard consumers (Excel/Word/etc) can preserve
 * basic cell formatting.
 *
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("../document/cell.js").CellState} CellState
 */

/**
 * @typedef {{ r: number, g: number, b: number }} RgbColor
 */

function clampByte(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(255, Math.round(n)));
}

/**
 * @param {string} value
 * @returns {RgbColor | null}
 */
function parseCssColorToRgb(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  if (trimmed.toLowerCase() === "transparent") return null;

  if (trimmed.startsWith("#")) {
    const hex = trimmed.slice(1);
    if (/^[0-9a-fA-F]{6}$/.test(hex)) {
      return {
        r: Number.parseInt(hex.slice(0, 2), 16),
        g: Number.parseInt(hex.slice(2, 4), 16),
        b: Number.parseInt(hex.slice(4, 6), 16),
      };
    }
    if (/^[0-9a-fA-F]{3}$/.test(hex)) {
      return {
        r: Number.parseInt(hex[0] + hex[0], 16),
        g: Number.parseInt(hex[1] + hex[1], 16),
        b: Number.parseInt(hex[2] + hex[2], 16),
      };
    }
    return null;
  }

  const rgbMatch = /^rgba?\((.*)\)$/i.exec(trimmed);
  if (rgbMatch) {
    const inner = rgbMatch[1] ?? "";
    const parts = inner
      .split(",")
      .map((p) => p.trim())
      .filter(Boolean);

    if (parts.length < 3) return null;

    const parseChannel = (raw) => {
      if (raw.endsWith("%")) {
        const pct = Number.parseFloat(raw.slice(0, -1));
        if (!Number.isFinite(pct)) return 0;
        return clampByte((pct / 100) * 255);
      }
      return clampByte(Number.parseFloat(raw));
    };

    let r = parseChannel(parts[0]);
    let g = parseChannel(parts[1]);
    let b = parseChannel(parts[2]);

    if (parts.length >= 4) {
      const a = Number.parseFloat(parts[3]);
      const alpha = Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1;
      // RTF doesn't support alpha. Approximate by blending with white.
      if (alpha < 1) {
        r = clampByte(alpha * r + (1 - alpha) * 255);
        g = clampByte(alpha * g + (1 - alpha) * 255);
        b = clampByte(alpha * b + (1 - alpha) * 255);
      }
    }

    return { r, g, b };
  }

  // Unknown CSS color (named colors, hsl(), etc). Leave unset.
  return null;
}

/**
 * @param {number} codePoint
 * @returns {number}
 */
function toRtfUnicodeValue(codePoint) {
  // RTF \u uses a signed 16-bit integer.
  const mod = codePoint % 0x10000;
  return mod > 0x7fff ? mod - 0x10000 : mod;
}

/**
 * Escape text for RTF.
 * - Escape `\`, `{`, `}`.
 * - Replace newlines with `\line`.
 * - Replace tabs with `\tab`.
 * - Emit basic unicode via `\uN?`.
 *
 * @param {string} text
 * @returns {string}
 */
function escapeRtfText(text) {
  const normalized = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");

  const escapeSegment = (segment) => {
    /** @type {string[]} */
    const out = [];
    for (const ch of segment) {
      if (ch === "\\") out.push("\\\\");
      else if (ch === "{") out.push("\\{");
      else if (ch === "}") out.push("\\}");
      else if (ch === "\t") out.push("\\tab ");
      else {
        const codePoint = ch.codePointAt(0) ?? 0;
        if (codePoint <= 0x7f) out.push(ch);
        else out.push(`\\u${toRtfUnicodeValue(codePoint)}?`);
      }
    }
    return out.join("");
  };

  return lines.map(escapeSegment).join("\\line ");
}

/**
 * RTF uses table indices starting at 1 (0 is the default/auto color).
 * @param {Map<string, number>} colorIndexByKey
 * @param {RgbColor[]} colors
 * @param {string | undefined} cssColor
 * @returns {number}
 */
function registerColor(colorIndexByKey, colors, cssColor) {
  const rgb = cssColor ? parseCssColorToRgb(cssColor) : null;
  if (!rgb) return 0;
  const key = `${rgb.r},${rgb.g},${rgb.b}`;
  const existing = colorIndexByKey.get(key);
  if (existing) return existing;
  colors.push(rgb);
  const index = colors.length; // 1-based; colors array doesn't include the leading default entry.
  colorIndexByKey.set(key, index);
  return index;
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellValueToRtf(cell) {
  const value = cell.value;
  if (value == null) {
    const formula = cell.formula;
    if (typeof formula === "string" && formula.trim() !== "") return formula;
    return "";
  }

  // DocumentController rich text values should copy as plain text.
  if (typeof value === "object" && typeof value.text === "string") return value.text;

  return String(value);
}

/**
 * Serialize a grid to a basic RTF table.
 *
 * @param {CellGrid} grid
 * @returns {string}
 */
export function serializeCellGridToRtf(grid) {
  const rows = Array.isArray(grid) ? grid : [];
  const colCount = rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);

  /** @type {Map<string, number>} */
  const colorIndexByKey = new Map();
  /** @type {RgbColor[]} */
  const colors = [];

  // Pre-scan for colors so we can build a single color table.
  for (const row of rows) {
    for (const cell of row ?? []) {
      const format = cell?.format;
      if (!format) continue;
      registerColor(colorIndexByKey, colors, format.textColor);
      registerColor(colorIndexByKey, colors, format.backgroundColor);
    }
  }

  const colorTable =
    `{\\colortbl;` +
    colors.map((c) => `\\red${c.r}\\green${c.g}\\blue${c.b};`).join("") +
    `}`;

  const CELL_WIDTH_TWIPS = 2000;

  /** @type {string[]} */
  const out = [];
  out.push("{\\rtf1\\ansi\\deff0");
  out.push("{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}}");
  out.push(colorTable);
  out.push("\\viewkind4\\uc1");

  for (const row of rows) {
    out.push("\\trowd\\trgaph108\\trleft0");

    // Cell descriptors.
    for (let col = 0; col < colCount; col++) {
      const cell = row?.[col] ?? { value: null, formula: null, format: null };
      const bg = cell?.format?.backgroundColor;
      const bgIndex = registerColor(colorIndexByKey, colors, bg);
      const shading = bgIndex > 0 ? 10000 : 0;
      const cellx = (col + 1) * CELL_WIDTH_TWIPS;
      out.push(`\\clcbpat${bgIndex}\\clshdng${shading}\\cellx${cellx}`);
    }

    // Cell contents.
    for (let col = 0; col < colCount; col++) {
      const cell = row?.[col] ?? { value: null, formula: null, format: null };
      const format = cell?.format ?? null;

      const textIndex = registerColor(colorIndexByKey, colors, format?.textColor);
      const bold = Boolean(format?.bold);
      const italic = Boolean(format?.italic);
      const underline = Boolean(format?.underline);

      const value = escapeRtfText(cellValueToRtf(cell));

      out.push("\\pard\\intbl");
      out.push(`\\cf${textIndex}`);
      out.push(bold ? "\\b" : "\\b0");
      out.push(italic ? "\\i" : "\\i0");
      out.push(underline ? "\\ul" : "\\ul0");
      out.push(` ${value}\\cell`);
    }

    out.push("\\row");
  }

  out.push("}");
  return out.join("\n");
}

