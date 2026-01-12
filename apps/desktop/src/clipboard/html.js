import { excelSerialToDate, parseScalar } from "../shared/valueParsing.js";

/**
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("../document/cell.js").CellState} CellState
 */

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function decodeHtmlEntities(text) {
  return String(text)
    .replaceAll("&nbsp;", " ")
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, code) => String.fromCharCode(Number.parseInt(code, 16)));
}

/**
 * Normalize Windows CF_HTML clipboard payloads to a plain HTML string that DOMParser can ingest.
 *
 * Windows "HTML Format" clipboard entries often look like:
 *   Version:0.9\r\nStartHTML:00000097\r\n...<html>...</html>
 *
 * When present, the numeric offsets are byte offsets into the full payload. In practice, they
 * can be missing or incorrect (especially with non-ASCII content), so we defensively fall back
 * to heuristics.
 *
 * @param {string} html
 * @returns {string}
 */
function normalizeClipboardHtml(html) {
  if (typeof html !== "string") return "";

  // Some producers include null padding/terminators.
  //
  // Important: strip trailing NULs eagerly (so we don't end up with odd characters at the end of
  // otherwise-valid markup), but avoid stripping *leading* NULs here. Some native clipboard bridges
  // include those leading bytes in CF_HTML offset calculations; removing them up-front would shift
  // StartHTML/StartFragment offsets and break extraction. Leading NULs are still removed later via
  // `stripToMarkup(...)`, which slices from the first HTML tag.
  const input = html.replace(/\u0000+$/, "");

  const findStartOfMarkup = (s) => {
    const doctype = s.search(/<!doctype/i);
    const htmlTag = s.search(/<html\b/i);
    const table = s.search(/<table\b/i);

    const candidates = [doctype, htmlTag, table].filter((i) => i >= 0);
    if (candidates.length === 0) return null;
    return Math.min(...candidates);
  };

  const stripToMarkup = (s) => {
    const idx = findStartOfMarkup(s);
    if (idx == null) return s;
    if (idx <= 0) return s;
    return s.slice(idx);
  };

  const headerEnd = findStartOfMarkup(input);
  const header = headerEnd == null ? input : input.slice(0, headerEnd);

  const getOffset = (name) => {
    const m = new RegExp(`(?:^|\\r?\\n)${name}:\\s*(-?\\d+)`, "i").exec(header);
    if (!m) return null;
    const n = Number.parseInt(m[1], 10);
    if (!Number.isFinite(n) || n < 0) return null;
    return n;
  };

  const startHtml = getOffset("StartHTML");
  const endHtml = getOffset("EndHTML");
  const startFragment = getOffset("StartFragment");
  const endFragment = getOffset("EndFragment");

  // CF_HTML offsets are byte offsets from the start of the payload. Prefer byte slicing
  // when possible, but fall back to code-unit slicing when needed.
  const encodeUtf8 = (text) => {
    if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(text);
    if (typeof Buffer !== "undefined") {
      // eslint-disable-next-line no-undef
      return Buffer.from(text, "utf8");
    }
    return null;
  };

  /** @type {Uint8Array | null | undefined} */
  let cachedUtf8Bytes;
  const getUtf8Bytes = () => {
    if (cachedUtf8Bytes !== undefined) return cachedUtf8Bytes;
    cachedUtf8Bytes = encodeUtf8(input);
    return cachedUtf8Bytes;
  };

  /** @type {TextDecoder | null | undefined} */
  let cachedUtf8Decoder;
  const getUtf8Decoder = () => {
    if (cachedUtf8Decoder !== undefined) return cachedUtf8Decoder;
    if (typeof TextDecoder === "undefined") {
      cachedUtf8Decoder = null;
      return cachedUtf8Decoder;
    }
    try {
      cachedUtf8Decoder = new TextDecoder("utf-8", { fatal: false });
    } catch {
      cachedUtf8Decoder = null;
    }
    return cachedUtf8Decoder;
  };

  const decodeUtf8 = (bytes) => {
    const decoder = getUtf8Decoder();
    if (decoder) return decoder.decode(bytes);
    if (typeof Buffer !== "undefined") {
      // eslint-disable-next-line no-undef
      return Buffer.from(bytes).toString("utf8");
    }
    return null;
  };

  const safeSliceUtf8 = (start, end) => {
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    if (start < 0 || end <= start) return null;
    const bytes = getUtf8Bytes();
    if (!bytes || end > bytes.length) return null;
    const decoded = decodeUtf8(bytes.subarray(start, end));
    return typeof decoded === "string" ? decoded : null;
  };

  const safeSliceCodeUnits = (start, end) => {
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    if (start < 0 || end <= start || end > input.length) return null;
    return input.slice(start, end);
  };

  const containsCompleteTable = (s) => /<table\b[\s\S]*?<\/table>/i.test(s);

  // Prefer fragment offsets when they look sane, but fall back to StartHTML/EndHTML.
  for (const [start, end] of [
    [startFragment, endFragment],
    [startHtml, endHtml],
  ]) {
    for (const candidate of [safeSliceUtf8(start, end), safeSliceCodeUnits(start, end)]) {
      if (!candidate) continue;
      const stripped = stripToMarkup(candidate);
      // Offsets can be "valid" but still wrong (e.g. truncated). Only accept them when they
      // contain a full table element so downstream parsing doesn't regress.
      if (containsCompleteTable(stripped)) return stripped;
    }
  }

  // If offsets are missing/incorrect, CF_HTML payloads often include fragment comment markers.
  // Use them as a best-effort way to isolate the correct table when multiple tables are present.
  const startMarker = /<!--\s*StartFragment\s*-->/i.exec(input);
  if (startMarker) {
    const afterStart = startMarker.index + startMarker[0].length;
    const endMarker = /<!--\s*EndFragment\s*-->/i.exec(input.slice(afterStart));
    if (endMarker) {
      const fragment = input.slice(afterStart, afterStart + endMarker.index);
      const stripped = stripToMarkup(fragment);
      if (containsCompleteTable(stripped)) return stripped;
    }
  }

  // Offsets missing or incorrect; fall back to heuristics on the full payload.
  return stripToMarkup(input);
}

function isLikelyDateNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  return fmt.toLowerCase().includes("yyyy-mm-dd");
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellValueToHtml(cell) {
  const value = cell.value;
  if (value == null) {
    const formula = cell.formula;
    if (typeof formula === "string" && formula.trim() !== "") {
      return escapeHtml(formula).replaceAll("\n", "<br>");
    }
    return "";
  }

  // DocumentController rich text values should copy as plain text.
  if (typeof value === "object" && typeof value.text === "string") {
    return escapeHtml(value.text).replaceAll("\n", "<br>");
  }

  const numberFormat = cell.format?.numberFormat;
  if (typeof value === "number" && isLikelyDateNumberFormat(numberFormat)) {
    const date = excelSerialToDate(value);
    const text = numberFormat.includes("hh") ? date.toISOString() : date.toISOString().slice(0, 10);
    return escapeHtml(text);
  }

  // Preserve embedded newlines via <br>.
  return escapeHtml(String(value)).replaceAll("\n", "<br>");
}

function formatToInlineStyle(format) {
  if (!format) return undefined;

  const rules = [];
  if (format.bold) rules.push("font-weight:bold");
  if (format.italic) rules.push("font-style:italic");
  if (format.underline) rules.push("text-decoration:underline");
  if (format.textColor) rules.push(`color:${format.textColor}`);
  if (format.backgroundColor) rules.push(`background-color:${format.backgroundColor}`);

  // Excel consumes `mso-number-format` from HTML clipboard payloads; keep it for round-tripping.
  if (format.numberFormat) rules.push(`mso-number-format:${format.numberFormat}`);

  return rules.length > 0 ? rules.join(";") : undefined;
}

function parseInlineStyle(style) {
  const rules = String(style)
    .split(";")
    .map((r) => r.trim())
    .filter(Boolean);

  const get = (name) => {
    const rule = rules.find((r) => r.toLowerCase().startsWith(`${name.toLowerCase()}:`));
    if (!rule) return undefined;
    return rule.slice(rule.indexOf(":") + 1).trim();
  };

  const fontWeight = get("font-weight");
  const bold =
    fontWeight !== undefined &&
    (fontWeight.toLowerCase() === "bold" || Number(fontWeight) >= 600);

  const italic = get("font-style")?.toLowerCase() === "italic";
  const textDecoration = get("text-decoration")?.toLowerCase() ?? "";
  const underline = textDecoration.includes("underline");

  const numberFormat = get("mso-number-format");
  const normalizedNumberFormat = numberFormat ? numberFormat.replace(/^['"]|['"]$/g, "") : undefined;

  const out = {};
  if (bold) out.bold = true;
  if (italic) out.italic = true;
  if (underline) out.underline = true;
  if (get("color")) out.textColor = get("color");
  if (get("background-color") ?? get("background")) {
    out.backgroundColor = get("background-color") ?? get("background");
  }
  if (normalizedNumberFormat) out.numberFormat = normalizedNumberFormat;
  return out;
}

function parseGoogleSheetsValue(data) {
  try {
    const parsed = JSON.parse(data);

    // Common patterns:
    // - {"1":2,"2":"text"} (string)
    // - {"1":3,"3":123} (number)
    if (typeof parsed?.["3"] === "number") return parsed["3"];
    if (typeof parsed?.["3"] === "boolean") return parsed["3"];
    if (typeof parsed?.["2"] === "string") return parsed["2"];

    const maybe = typeof parsed?.["3"] === "string" ? parsed["3"] : typeof parsed?.["2"] === "string" ? parsed["2"] : undefined;
    if (typeof maybe === "string") return maybe;
  } catch {
    // Ignore; fall back to text parsing.
  }
  return undefined;
}

/**
 * Extract plain text from a DOM table cell while preserving explicit line breaks.
 *
 * DOMParser's `textContent` drops `<br>` elements, collapsing multiline content.
 * This walker reconstitutes those line breaks as `\n` for proper round-tripping.
 *
 * @param {Element} cellEl
 * @returns {string}
 */
function extractCellTextDom(cellEl) {
  /** @type {string[]} */
  const out = [];

  /**
   * @param {Node} node
   */
  const walk = (node) => {
    // Text node.
    if (node.nodeType === 3) {
      out.push(node.nodeValue ?? "");
      return;
    }

    // Element node.
    if (node.nodeType === 1) {
      const el = /** @type {Element} */ (node);
      if (el.tagName.toLowerCase() === "br") {
        out.push("\n");
        return;
      }

      for (const child of Array.from(el.childNodes)) walk(child);
    }
  };

  for (const child of Array.from(cellEl.childNodes)) walk(child);

  return out.join("").replaceAll("\u00a0", " ");
}

/**
 * @param {CellGrid} grid
 * @returns {string}
 */
export function serializeCellGridToHtml(grid) {
  const rows = grid
    .map((row) => {
      const tds = row
        .map((cell) => {
          const style = formatToInlineStyle(cell.format);
          const styleAttr = style ? ` style="${escapeHtml(style)}"` : "";
          const formulaAttr = cell.formula
            ? ` data-formula="${escapeHtml(cell.formula)}" data-sheets-formula="${escapeHtml(
                cell.formula,
              )}" x:formula="${escapeHtml(cell.formula)}"`
            : "";
          const numberFormatAttr =
            cell.format?.numberFormat != null
              ? ` data-number-format="${escapeHtml(cell.format.numberFormat)}"`
              : "";
          return `<td${styleAttr}${formulaAttr}${numberFormatAttr}>${cellValueToHtml(cell)}</td>`;
        })
        .join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");

  // Many clipboard consumers (including Excel) look for this fragment wrapper.
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><!--StartFragment--><table>${rows}</table><!--EndFragment--></body></html>`;
}

/**
 * @param {string} html
 * @returns {CellGrid | null}
 */
export function parseHtmlToCellGrid(html) {
  const normalized = normalizeClipboardHtml(html);
  if (typeof DOMParser !== "undefined") return parseHtmlToCellGridDom(normalized);
  return parseHtmlToCellGridFallback(normalized);
}

/**
 * DOM-based parser (browser / WebView).
 * @param {string} html
 * @returns {CellGrid | null}
 */
function parseHtmlToCellGridDom(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;

  /** @type {CellGrid} */
  const grid = [];

  for (const row of Array.from(table.querySelectorAll("tr"))) {
    const cells = Array.from(row.querySelectorAll("th,td"));
    /** @type {CellState[]} */
    const outRow = [];

    for (const cellEl of cells) {
      const formula =
        cellEl.getAttribute("data-formula") ??
        cellEl.getAttribute("data-sheets-formula") ??
        cellEl.getAttribute("x:formula") ??
        null;

      const style = cellEl.getAttribute("style");
      const parsedStyle = style ? parseInlineStyle(style) : {};

      const numberFormatAttr = cellEl.getAttribute("data-number-format");
      if (numberFormatAttr) parsedStyle.numberFormat ??= numberFormatAttr;

      const img = cellEl.querySelector("img");
      if (img) {
        const alt = img.getAttribute("alt") ?? "image";
        outRow.push({ value: alt, formula, format: Object.keys(parsedStyle).length ? parsedStyle : null });
        continue;
      }

      const sheetsValueAttr = cellEl.getAttribute("data-sheets-value");
      const excelNumAttr = cellEl.getAttribute("x:num");

      let raw;
      if (sheetsValueAttr) raw = parseGoogleSheetsValue(sheetsValueAttr);
      if (raw === undefined && excelNumAttr) raw = excelNumAttr;
      if (raw === undefined) raw = extractCellTextDom(cellEl);

      const parsed = parseScalar(String(raw));
      if (parsed.type === "datetime" && !parsedStyle.numberFormat) {
        parsedStyle.numberFormat = parsed.numberFormat;
      }

      outRow.push({
        value: parsed.value,
        formula,
        format: Object.keys(parsedStyle).length ? parsedStyle : null,
      });
    }

    grid.push(outRow);
  }

  return grid;
}

/**
 * Regex-based fallback parser for non-DOM environments (Node tests).
 * @param {string} html
 * @returns {CellGrid | null}
 */
function parseHtmlToCellGridFallback(html) {
  const tableMatch = /<table\b[\s\S]*?<\/table>/i.exec(html);
  if (!tableMatch) return null;

  const tableHtml = tableMatch[0];
  const rowRegex = /<tr\b[\s\S]*?<\/tr>/gi;
  const cellRegex = /<(td|th)\b([^>]*)>([\s\S]*?)<\/\1>/gi;

  /** @type {CellGrid} */
  const grid = [];

  for (const rowHtml of tableHtml.match(rowRegex) ?? []) {
    /** @type {CellState[]} */
    const row = [];
    for (const cellMatch of rowHtml.matchAll(cellRegex)) {
      const attrs = cellMatch[2] ?? "";
      const inner = cellMatch[3] ?? "";

      const getAttr = (name) => {
        const re = new RegExp(`${name}\\s*=\\s*(\"([^\"]*)\"|'([^']*)'|([^\\s>]+))`, "i");
        const m = re.exec(attrs);
        if (!m) return undefined;
        return decodeHtmlEntities(m[2] ?? m[3] ?? m[4] ?? "");
      };

      const formula = getAttr("data-formula") ?? getAttr("data-sheets-formula") ?? getAttr("x:formula") ?? null;

      const style = getAttr("style");
      const parsedStyle = style ? parseInlineStyle(style) : {};
      const numberFormatAttr = getAttr("data-number-format");
      if (numberFormatAttr) parsedStyle.numberFormat ??= numberFormatAttr;

      // Image placeholder.
      const imgAlt = /<img\b[^>]*\balt\s*=\s*(\"([^\"]*)\"|'([^']*)')/i.exec(inner);
      if (imgAlt) {
        const alt = decodeHtmlEntities(imgAlt[2] ?? imgAlt[3] ?? "image");
        row.push({ value: alt, formula, format: Object.keys(parsedStyle).length ? parsedStyle : null });
        continue;
      }
      if (/<img\b/i.test(inner)) {
        row.push({ value: "image", formula, format: Object.keys(parsedStyle).length ? parsedStyle : null });
        continue;
      }

      const sheetsValueAttr = getAttr("data-sheets-value");
      const excelNumAttr = getAttr("x:num");

      let raw;
      if (sheetsValueAttr) raw = parseGoogleSheetsValue(sheetsValueAttr);
      if (raw === undefined && excelNumAttr) raw = excelNumAttr;
      if (raw === undefined) {
        raw = decodeHtmlEntities(
          inner
            .replace(/<!--[\s\S]*?-->/g, "")
            .replace(/<br\s*\/?>/gi, "\n")
            .replace(/<[^>]+>/g, "")
        ).replaceAll("\u00a0", " ");
      }

      const parsed = parseScalar(String(raw));
      if (parsed.type === "datetime" && !parsedStyle.numberFormat) {
        parsedStyle.numberFormat = parsed.numberFormat;
      }

      row.push({
        value: parsed.value,
        formula,
        format: Object.keys(parsedStyle).length ? parsedStyle : null,
      });
    }

    grid.push(row);
  }

  return grid;
}
