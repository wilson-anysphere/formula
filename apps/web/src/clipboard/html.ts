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

