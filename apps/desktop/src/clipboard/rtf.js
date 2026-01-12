/**
 * RTF clipboard helpers.
 *
 * - `serializeCellGridToRtf`: emit a basic RTF table (for rich clipboard consumers).
 * - `extractPlainTextFromRtf`: best-effort RTF -> plain text extraction (for paste fallback).
 *
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("../document/cell.js").CellState} CellState
 */

/**
 * @typedef {{ r: number, g: number, b: number }} RgbColor
 */

/**
 * Minimal CSS named color support so clipboard formats that use named colors
 * (e.g. "red", "yellow", theme values like "rebeccapurple") serialize into a
 * deterministic RTF color table even in non-DOM environments (Node tests).
 *
 * This is not an exhaustive list, but covers common spreadsheet colors and
 * theme tokens used in this codebase.
 *
 * @type {Record<string, RgbColor>}
 */
const CSS_NAMED_COLORS = {
  // Basic HTML colors + common aliases.
  black: { r: 0, g: 0, b: 0 },
  silver: { r: 192, g: 192, b: 192 },
  gray: { r: 128, g: 128, b: 128 },
  grey: { r: 128, g: 128, b: 128 },
  white: { r: 255, g: 255, b: 255 },
  maroon: { r: 128, g: 0, b: 0 },
  red: { r: 255, g: 0, b: 0 },
  purple: { r: 128, g: 0, b: 128 },
  fuchsia: { r: 255, g: 0, b: 255 },
  magenta: { r: 255, g: 0, b: 255 },
  green: { r: 0, g: 128, b: 0 },
  lime: { r: 0, g: 255, b: 0 },
  olive: { r: 128, g: 128, b: 0 },
  yellow: { r: 255, g: 255, b: 0 },
  navy: { r: 0, g: 0, b: 128 },
  blue: { r: 0, g: 0, b: 255 },
  teal: { r: 0, g: 128, b: 128 },
  aqua: { r: 0, g: 255, b: 255 },
  cyan: { r: 0, g: 255, b: 255 },

  // Common extras.
  orange: { r: 255, g: 165, b: 0 },
  hotpink: { r: 255, g: 105, b: 180 },
  rebeccapurple: { r: 102, g: 51, b: 153 },
};

function clampByte(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(255, Math.round(n)));
}

/**
 * @param {string} value
 * @returns {string | null}
 */
function normalizeCssColorViaDom(value) {
  try {
    // eslint-disable-next-line no-undef
    const doc = typeof document !== "undefined" ? document : null;
    // eslint-disable-next-line no-undef
    const compute = typeof getComputedStyle === "function" ? getComputedStyle : null;
    if (!doc || typeof doc.createElement !== "function" || !compute) return null;

    const el = doc.createElement("span");
    el.style.color = "";
    el.style.color = value;
    if (!el.style.color) return null;

    const parent = doc.body ?? doc.documentElement;
    if (parent && typeof parent.appendChild === "function") parent.appendChild(el);
    const computed = compute(el).color;
    el.remove();
    return typeof computed === "string" && computed.trim() ? computed.trim() : null;
  } catch {
    return null;
  }
}

/**
 * @param {string} value
 * @returns {RgbColor | null}
 */
function parseCssColorToRgb(value) {
  const parsed = parseCssColorToRgbNoDom(value);
  if (parsed) return parsed;

  const normalized = normalizeCssColorViaDom(value);
  if (!normalized) return null;
  return parseCssColorToRgbNoDom(normalized);
}

/**
 * @param {string} value
 * @returns {RgbColor | null}
 */
function parseCssColorToRgbNoDom(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  const lower = trimmed.toLowerCase();
  if (lower === "transparent" || lower === "none") return null;

  const named = CSS_NAMED_COLORS[lower];
  if (named) return named;

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

  const match = /^(rgb|rgba)\(\s*([\s\S]+)\s*\)$/i.exec(trimmed);
  if (match) {
    let args = match[2]?.trim() ?? "";
    if (!args) return null;

    let alphaPart = null;

    // Support modern slash syntax: rgb(… / …).
    if (args.includes("/")) {
      const parts = args.split("/");
      if (parts.length !== 2) return null;
      args = parts[0].trim();
      alphaPart = parts[1].trim();
    }

    let parts = [];
    if (args.includes(",")) {
      parts = args
        .split(",")
        .map((p) => p.trim())
        .filter(Boolean);
    } else {
      parts = args
        .split(/\s+/)
        .map((p) => p.trim())
        .filter(Boolean);
    }

    if (alphaPart == null && parts.length === 4) {
      alphaPart = parts[3];
      parts = parts.slice(0, 3);
    }

    if (parts.length < 3) return null;

    const parseChannel = (raw) => {
      if (raw.endsWith("%")) {
        const pct = Number.parseFloat(raw.slice(0, -1));
        if (!Number.isFinite(pct)) return 0;
        return clampByte((pct / 100) * 255);
      }
      return clampByte(Number.parseFloat(raw));
    };

    const r = parseChannel(parts[0]);
    const g = parseChannel(parts[1]);
    const b = parseChannel(parts[2]);

    if (alphaPart == null) return { r, g, b };

    let alpha;
    if (alphaPart.endsWith("%")) {
      const pct = Number.parseFloat(alphaPart.slice(0, -1));
      alpha = Number.isFinite(pct) ? pct / 100 : 1;
    } else {
      alpha = Number.parseFloat(alphaPart);
    }
    alpha = Number.isFinite(alpha) ? Math.max(0, Math.min(1, alpha)) : 1;

    // RTF doesn't support alpha. Approximate by blending with white.
    if (alpha >= 1) return { r, g, b };
    return {
      r: clampByte(alpha * r + (1 - alpha) * 255),
      g: clampByte(alpha * g + (1 - alpha) * 255),
      b: clampByte(alpha * b + (1 - alpha) * 255),
    };
  }

  // Unknown CSS color (hsl(), var(--foo), etc). Leave unset.
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
 * - Escape `\\`, `{`, `}`.
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
    // RTF `\uN` escape values are UTF-16 code units (signed 16-bit ints), not
    // full Unicode code points. Iterate by code unit so astral-plane characters
    // (surrogate pairs) round-trip correctly.
    for (let i = 0; i < segment.length; i++) {
      const ch = segment[i];
      if (ch === "\\") out.push("\\\\");
      else if (ch === "{") out.push("\\{");
      else if (ch === "}") out.push("\\}");
      else if (ch === "\t") out.push("\\tab ");
      else {
        const codeUnit = segment.charCodeAt(i);
        if (codeUnit <= 0x7f) out.push(ch);
        else out.push(`\\u${toRtfUnicodeValue(codeUnit)}?`);
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
 * @param {Map<string, number>} colorIndexByKey
 * @param {string | undefined} cssColor
 * @returns {number}
 */
function getColorIndex(colorIndexByKey, cssColor) {
  const rgb = cssColor ? parseCssColorToRgb(cssColor) : null;
  if (!rgb) return 0;
  const key = `${rgb.r},${rgb.g},${rgb.b}`;
  return colorIndexByKey.get(key) ?? 0;
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
      const bgIndex = getColorIndex(colorIndexByKey, bg);
      const shading = bgIndex > 0 ? 10000 : 0;
      const cellx = (col + 1) * CELL_WIDTH_TWIPS;
      out.push(`\\clcbpat${bgIndex}\\clshdng${shading}\\cellx${cellx}`);
    }

    // Cell contents.
    for (let col = 0; col < colCount; col++) {
      const cell = row?.[col] ?? { value: null, formula: null, format: null };
      const format = cell?.format ?? null;

      const textIndex = getColorIndex(colorIndexByKey, format?.textColor);
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

/**
 * Best-effort RTF -> plain text extraction.
 *
 * The clipboard can sometimes expose only `text/rtf` without `text/plain` or `text/html`
 * (e.g. when using OS-native clipboard backends). We only need enough fidelity to
 * recover tabular content for TSV parsing.
 *
 * This is intentionally conservative: we translate a small set of common control
 * words and drop everything else (control words, groups, destinations).
 *
 * @param {string} rtf
 * @returns {string}
 */
export function extractPlainTextFromRtf(rtf) {
  if (typeof rtf !== "string" || rtf.trim() === "") return "";

  /** @type {Set<string>} */
  const IGNORED_DESTINATIONS = new Set([
    "fonttbl",
    "colortbl",
    "stylesheet",
    "info",
    "pict",
    "object",
    "datastore",
    "themedata",
  ]);

  let out = "";

  // RTF state is scoped to groups; treat `\ucN` as state.
  let ignorable = false;
  let ucSkip = 1;
  let atStart = true;
  /** @type {{ ignorable: boolean, ucSkip: number, atStart: boolean }[]} */
  const stack = [];

  const len = rtf.length;
  let i = 0;

  while (i < len) {
    const ch = rtf[i];

    if (ch === "{") {
      stack.push({ ignorable, ucSkip, atStart });
      atStart = true;
      i += 1;
      continue;
    }

    if (ch === "}") {
      const prev = stack.pop();
      if (prev) {
        ignorable = prev.ignorable;
        ucSkip = prev.ucSkip;
        atStart = prev.atStart;
      }
      i += 1;
      continue;
    }

    // Newlines inside RTF are usually formatting artifacts, not literal content.
    if (ch === "\r" || ch === "\n") {
      i += 1;
      continue;
    }

    if (ch !== "\\") {
      if (!ignorable) out += ch;
      atStart = false;
      i += 1;
      continue;
    }

    // Backslash-initiated control sequence.
    if (i + 1 >= len) break;
    const next = rtf[i + 1];

    // Escaped literal.
    if (next === "\\" || next === "{" || next === "}") {
      if (!ignorable) out += next;
      atStart = false;
      i += 2;
      continue;
    }

    // Hex-escaped character: \'hh
    if (next === "'") {
      const hex = rtf.slice(i + 2, i + 4);
      if (/^[0-9a-fA-F]{2}$/.test(hex)) {
        if (!ignorable) out += String.fromCharCode(Number.parseInt(hex, 16));
        atStart = false;
        i += 4;
        continue;
      }
      // Malformed; drop the escape introducer.
      i += 2;
      continue;
    }

    // Control symbols are backslash + single non-letter.
    if (!/[a-zA-Z]/.test(next)) {
      if (!ignorable) {
        if (next === "~") out += " ";
        else if (next === "_") out += "-";
        else if (next === "-") out += ""; // optional hyphen
        // Else: drop unknown control symbols.
      }

      if (next === "*") {
        // Destination marked as ignorable.
        ignorable = true;
      }

      atStart = false;
      i += 2;
      continue;
    }

    // Control word: \wordN?
    let j = i + 1;
    while (j < len && /[a-zA-Z]/.test(rtf[j])) j += 1;
    const word = rtf.slice(i + 1, j);

    // Optional numeric parameter.
    let sign = 1;
    if (rtf[j] === "-") {
      sign = -1;
      j += 1;
    } else if (rtf[j] === "+") {
      j += 1;
    }

    let numStr = "";
    while (j < len && /[0-9]/.test(rtf[j])) {
      numStr += rtf[j];
      j += 1;
    }

    const param = numStr ? sign * Number.parseInt(numStr, 10) : null;

    // A space following a control word is a delimiter and should be consumed.
    if (rtf[j] === " ") j += 1;

    // Some destinations should be skipped entirely (font tables, embedded images, etc).
    if (atStart) {
      if (IGNORED_DESTINATIONS.has(word)) ignorable = true;
      atStart = false;
    }

    if (!ignorable) {
      if (word === "tab") {
        out += "\t";
      } else if (word === "par" || word === "line") {
        out += "\n";
      } else if (word === "u" && typeof param === "number") {
        // `\uN` is a signed 16-bit integer; map negatives back into [0, 65535].
        let code = param;
        if (code < 0) code = 65536 + code;
        out += String.fromCharCode(code);
      } else if (word === "uc" && typeof param === "number") {
        // Number of "fallback" characters to skip after a unicode escape.
        ucSkip = Math.max(0, param);
      }
    } else if (word === "uc" && typeof param === "number") {
      // Even in ignored destinations, maintain `\ucN` state so `\uN` skipping stays in sync.
      ucSkip = Math.max(0, param);
    }

    i = j;

    // After \uN, skip the ANSI fallback characters (usually a single '?'), which are
    // not meaningful once we've emitted the unicode value.
    if (word === "u" && typeof param === "number" && ucSkip > 0) {
      let skipped = 0;
      while (i < len && skipped < ucSkip) {
        const c = rtf[i];
        if (c === "{" || c === "}") break;
        if (c === "\\") {
          // Count a hex escape as one fallback char, otherwise stop before control words.
          if (rtf[i + 1] === "'" && /^[0-9a-fA-F]{2}$/.test(rtf.slice(i + 2, i + 4))) {
            i += 4;
            skipped += 1;
            continue;
          }
          break;
        }
        i += 1;
        skipped += 1;
      }
    }
  }

  return out;
}
