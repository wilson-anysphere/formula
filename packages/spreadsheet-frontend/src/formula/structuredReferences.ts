export type ParsedStructuredReference = {
  tableName: string;
  columnName: string;
  includeHeader: boolean;
};

export function parseStructuredReferenceText(text: string): ParsedStructuredReference | null {
  const firstBracket = text.indexOf("[");
  if (firstBracket <= 0) return null;
  const tableName = text.slice(0, firstBracket);
  const suffix = text.slice(firstBracket);
  if (!tableName || !suffix) return null;

  // Supported patterns:
  //   TableName[ColumnName]
  //   TableName[[#All],[ColumnName]]
  const allMatch = /^\[\[\s*#all\s*\]\s*,\s*\[\s*([^\]]+?)\s*\]\]$/i.exec(suffix);
  if (allMatch) {
    return { tableName, columnName: allMatch[1]!, includeHeader: true };
  }

  const simpleMatch = /^\[\s*([^\[\]]+?)\s*\]$/.exec(suffix);
  if (simpleMatch) {
    return { tableName, columnName: simpleMatch[1]!, includeHeader: false };
  }

  return null;
}

