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

const ESCAPED_ITEM_RE_SRC = "((?:[^\\]]|\\]\\])+)"; // match non-] or escaped `]]`
const QUALIFIED_REFERENCE_RE = new RegExp(
  `^\\[\\[\\s*${ESCAPED_ITEM_RE_SRC}\\s*\\]\\s*,\\s*\\[\\s*${ESCAPED_ITEM_RE_SRC}\\s*\\]\\]$`,
  "i"
);

// Excel permits "this row" shorthand forms:
//   TableName[@Column]
//   TableName[@[Column Name]]
//   TableName[@]               (whole-row this-row reference)
// and, within the table context, the implicit equivalents:
//   [@Column]
//   [@[Column Name]]
//   [@]                        (whole-row this-row reference)
//
// The nested-bracket form exists because column names can contain spaces/special chars.
// Column names may still contain escaped closing brackets via doubling: `]` -> `]]`.
const THIS_ROW_NESTED_REFERENCE_RE = new RegExp(
  `^\\[\\s*@\\s*\\[\\s*${ESCAPED_ITEM_RE_SRC}\\s*\\]\\s*\\]$`,
  "i"
);
const SIMPLE_REFERENCE_RE = new RegExp(`^\\[\\s*${ESCAPED_ITEM_RE_SRC}\\s*\\]$`);

function unescapeStructuredRefItem(text: string): string {
  // Excel escapes `]` inside structured reference items by doubling it: `]]` -> `]`.
  return text.replaceAll("]]", "]");
}

export function parseStructuredReferenceText(text: string): ParsedStructuredReference | null {
  const firstBracket = text.indexOf("[");
  if (firstBracket < 0) return null;
  const tableName = text.slice(0, firstBracket);
  const suffix = text.slice(firstBracket);
  if (!suffix) return null;

  // Avoid treating arbitrary `[X]` bracket groups as structured references when the
  // table prefix is omitted. Today we only recognize:
  // - implicit this-row references (`[@...]`)
  // - implicit selector-qualified references (`[[#...],[...]]`)
  // This prevents tokenizing fragments like `[Book.xlsx]` (external workbook refs) as tables.
  if (!tableName) {
    const trimmed = suffix.trimStart();
    if (!(trimmed.startsWith("[@") || trimmed.startsWith("[["))) return null;
  }

  // Supported patterns:
  //   TableName[ColumnName]
  //   TableName[#All]
  //   TableName[@ColumnName]
  //   TableName[@[Column Name]]
  //   TableName[@] (entire row; columnName will be empty)
  //   TableName[[#All],[ColumnName]]
  //   TableName[[#Data],[ColumnName]]
  //   TableName[[#Headers],[ColumnName]]
  //
  // Column names may include escaped closing brackets: `]` is encoded as `]]`.
  const qualifiedMatch = QUALIFIED_REFERENCE_RE.exec(suffix);
  if (qualifiedMatch) {
    const selector = unescapeStructuredRefItem(qualifiedMatch[1]!.trim());
    if (!selector.startsWith("#")) return null;
    const columnName = unescapeStructuredRefItem(qualifiedMatch[2]!.trim());
    // Ambiguous: `Table1[[#All],[#Totals]]` (and friends) represent *multiple* structured
    // reference items with no explicit column selection (all columns). These should be handled
    // by a nested structured-ref parser rather than being misinterpreted as a selector-qualified
    // single-item reference where the "column name" is itself a selector.
    const normalizedColumnItem = columnName.trim().replace(/\s+/g, " ").toLowerCase();
    if (
      normalizedColumnItem === "#all" ||
      normalizedColumnItem === "#headers" ||
      normalizedColumnItem === "#data" ||
      normalizedColumnItem === "#totals" ||
      normalizedColumnItem === "#this row"
    ) {
      return null;
    }
    return { tableName, selector, columnName };
  }

  // `TableName[@[Column Name]]` (or implicit `[@[Column Name]]`).
  const thisRowNested = THIS_ROW_NESTED_REFERENCE_RE.exec(suffix);
  if (thisRowNested) {
    const columnName = unescapeStructuredRefItem(thisRowNested[1]!.trim());
    return { tableName, selector: "#This Row", columnName };
  }

  // Avoid mis-parsing nested bracket groups like `[[#All],[Amount]]` as a single item.
  if (suffix.startsWith("[[")) return null;

  const simpleMatch = SIMPLE_REFERENCE_RE.exec(suffix);
  if (simpleMatch) {
    const item = unescapeStructuredRefItem(simpleMatch[1]!.trim());
    // `TableName[@Column]` / `TableName[@]` (or implicit `[@...]`) is shorthand for `#This Row`.
    if (item.startsWith("@")) {
      const columnName = item.slice(1).trim();
      // Whole-row shorthand: `Table[@]` / `[@]` refers to the current row's values.
      if (!columnName) return { tableName, selector: "#This Row", columnName: "" };
      // Shorthand `@Column` does not permit whitespace or nested bracket syntax; column names
      // that require quoting (spaces, etc) use the `@[[Column Name]]` form handled above.
      // Keep this strict to avoid accidentally treating `[@[Col]] , ...` tails as a single token
      // when `]]` sequences appear (which are ambiguous between escapes and nested closes).
      if (columnName && /[\s\[\],;\]]/.test(columnName)) return null;
      return { tableName, selector: "#This Row", columnName };
    }
    return { tableName, selector: null, columnName: item };
  }

  return null;
}
