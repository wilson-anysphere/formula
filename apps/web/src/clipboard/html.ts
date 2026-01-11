function escapeHtml(text: string): string {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
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

export function parseHtmlTableToGrid(html: string): string[][] | null {
  if (typeof DOMParser === "undefined") return null;

  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return null;

  const rows: string[][] = [];
  for (const row of Array.from(table.querySelectorAll("tr"))) {
    const cells = Array.from(row.querySelectorAll("th,td")).map((cell) =>
      nodeToText(cell).replaceAll("\u00a0", " ")
    );
    rows.push(cells);
  }

  return rows;
}
