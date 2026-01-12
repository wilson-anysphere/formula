function escapeHtml(text: string): string {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function decodeHtmlEntities(text: string): string {
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

export function serializeGridToHtmlTable(grid: string[][]): string {
  const rows = grid
    .map((row) => {
      const cells = row
        .map((value) => {
          const escaped = escapeHtml(value).replaceAll("\n", "<br>");
          return `<td>${escaped}</td>`;
        })
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  // Many clipboard consumers (including Excel) look for this fragment wrapper.
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><!--StartFragment--><table>${rows}</table><!--EndFragment--></body></html>`;
}

function nodeToText(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent ?? "";
  if (node.nodeType !== Node.ELEMENT_NODE) return "";

  const el = node as Element;
  if (el.tagName.toLowerCase() === "br") return "\n";

  let text = "";
  for (const child of Array.from(el.childNodes)) {
    text += nodeToText(child);
  }
  return text;
}

export type ParseGridOptions = {
  /**
   * Hard cap on the number of parsed cells.
   *
   * HTML table payloads can represent Excel-scale selections (millions of cells). Parsing them into
   * a full 2D JS array can easily OOM the tab/renderer.
   */
  maxCells?: number;
  maxRows?: number;
  maxCols?: number;
  /**
   * Hard cap on the HTML payload size. Large HTML tables can cause DOMParser to allocate huge
   * intermediate structures; reject them early.
   */
  maxChars?: number;
};

const DEFAULT_MAX_CLIPBOARD_CELLS = 200_000;
const DEFAULT_MAX_CLIPBOARD_CHARS = 10_000_000;

function parseHtmlTableToGridDom(html: string, options: ParseGridOptions): string[][] | null {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;

  const rows: string[][] = [];
  const maxCells = options.maxCells ?? DEFAULT_MAX_CLIPBOARD_CELLS;
  const maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
  const maxCols = options.maxCols ?? Number.POSITIVE_INFINITY;
  let cellCount = 0;

  for (const row of Array.from(table.querySelectorAll("tr"))) {
    if (rows.length >= maxRows) return null;
    const rawCells = Array.from(row.querySelectorAll("th,td"));
    if (rawCells.length > maxCols) return null;
    const cells = rawCells.map((cell) => nodeToText(cell).replaceAll("\u00a0", " "));
    cellCount += cells.length;
    if (cellCount > maxCells) return null;
    rows.push(cells);
  }

  return rows;
}

function parseHtmlTableToGridFallback(html: string, options: ParseGridOptions): string[][] | null {
  const tableMatch = /<table\b[\s\S]*?<\/table>/i.exec(html);
  if (!tableMatch) return null;

  const tableHtml = tableMatch[0];
  const rowRegex = /<tr\b[\s\S]*?<\/tr>/gi;
  const cellRegex = /<(td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi;

  const grid: string[][] = [];
  const maxCells = options.maxCells ?? DEFAULT_MAX_CLIPBOARD_CELLS;
  const maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
  const maxCols = options.maxCols ?? Number.POSITIVE_INFINITY;
  let cellCount = 0;

  for (const rowHtml of tableHtml.match(rowRegex) ?? []) {
    if (grid.length >= maxRows) return null;
    const row: string[] = [];
    for (const cellMatch of rowHtml.matchAll(cellRegex)) {
      if (row.length >= maxCols) return null;
      const inner = cellMatch[2] ?? "";
      const value = decodeHtmlEntities(
        inner
          .replace(/<!--[\s\S]*?-->/g, "")
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "")
      ).replaceAll("\u00a0", " ");
      row.push(value);
      cellCount += 1;
      if (cellCount > maxCells) return null;
    }
    grid.push(row);
  }

  return grid;
}

export function parseHtmlTableToGrid(html: string, options: ParseGridOptions = {}): string[][] | null {
  const maxChars = options.maxChars ?? DEFAULT_MAX_CLIPBOARD_CHARS;
  const text = String(html ?? "");
  if (text.length > maxChars) return null;
  if (typeof DOMParser !== "undefined") return parseHtmlTableToGridDom(text, options);
  return parseHtmlTableToGridFallback(text, options);
}
