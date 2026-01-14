export type ParsedStructuredReference = {
  tableName: string;
  columnName: string;
  /**
   * Optional structured reference selector (e.g. `#All`, `#Data`, `#Headers`).
   *
   * When present, this corresponds to `Table1[[#Selector],[Column]]`-style references.
   */
  selector: string | null;
};

function unescapeStructuredRefItem(text: string): string {
  // Excel escapes `]` inside structured reference items by doubling it: `]]` -> `]`.
  return text.replaceAll("]]", "]");
}

export function parseStructuredReferenceText(text: string): ParsedStructuredReference | null {
  const firstBracket = text.indexOf("[");
  if (firstBracket <= 0) return null;
  const tableName = text.slice(0, firstBracket);
  const suffix = text.slice(firstBracket);
  if (!tableName || !suffix) return null;

  // Supported patterns:
  //   TableName[ColumnName]
  //   TableName[#All]
  //   TableName[[#All],[ColumnName]]
  //   TableName[[#Data],[ColumnName]]
  //   TableName[[#Headers],[ColumnName]]
  //
  // Column names may include escaped closing brackets: `]` is encoded as `]]`.
  const escapedItem = "((?:[^\\]]|\\]\\])+)"; // match non-] or escaped `]]`
  const qualifiedRe = new RegExp(`^\\[\\[\\s*${escapedItem}\\s*\\]\\s*,\\s*\\[\\s*${escapedItem}\\s*\\]\\]$`, "i");
  const qualifiedMatch = qualifiedRe.exec(suffix);
  if (qualifiedMatch) {
    const selector = unescapeStructuredRefItem(qualifiedMatch[1]!.trim());
    if (!selector.startsWith("#")) return null;
    const columnName = unescapeStructuredRefItem(qualifiedMatch[2]!.trim());
    return { tableName, selector, columnName };
  }

  // Avoid mis-parsing nested bracket groups like `[[#All],[Amount]]` as a single item.
  if (suffix.startsWith("[[")) return null;

  const simpleRe = new RegExp(`^\\[\\s*${escapedItem}\\s*\\]$`);
  const simpleMatch = simpleRe.exec(suffix);
  if (simpleMatch) {
    return { tableName, selector: null, columnName: unescapeStructuredRefItem(simpleMatch[1]!.trim()) };
  }

  return null;
}
