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
  if (typeof DOMParser !== "undefined") return parseHtmlToCellGridDom(html);
  return parseHtmlToCellGridFallback(html);
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
      if (raw === undefined) raw = (cellEl.textContent ?? "").replaceAll("\u00a0", " ");

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
