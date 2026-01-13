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
  const qualifiedMatch = /^\[\[\s*(#[A-Za-z]+)\s*\]\s*,\s*\[\s*([^\]]+?)\s*\]\]$/i.exec(suffix);
  if (qualifiedMatch) {
    return { tableName, selector: qualifiedMatch[1]!, columnName: qualifiedMatch[2]! };
  }

  const simpleMatch = /^\[\s*([^\[\]]+?)\s*\]$/.exec(suffix);
  if (simpleMatch) {
    return { tableName, selector: null, columnName: simpleMatch[1]! };
  }

  return null;
}
