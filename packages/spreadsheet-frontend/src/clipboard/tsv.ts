export function serializeGridToTsv(grid: string[][]): string {
  return grid.map((row) => row.join("\t")).join("\n");
}

export type ParseGridOptions = {
  /**
   * Hard cap on the number of cells to parse.
   *
   * Clipboard TSV payloads can represent Excel-scale selections (millions of cells). Parsing
   * them into a full 2D JS array can easily OOM the tab/renderer.
   */
  maxCells?: number;
  maxRows?: number;
  maxCols?: number;
};

const DEFAULT_MAX_CLIPBOARD_CELLS = 200_000;

export function parseTsvToGrid(tsv: string, options: ParseGridOptions = {}): string[][] | null {
  const maxCells = options.maxCells ?? DEFAULT_MAX_CLIPBOARD_CELLS;
  const maxRows = options.maxRows ?? Number.POSITIVE_INFINITY;
  const maxCols = options.maxCols ?? Number.POSITIVE_INFINITY;

  const text = String(tsv ?? "");
  /** @type {string[][]} */
  const grid = [];
  /** @type {string[]} */
  let row = [];

  let cellStart = 0;
  let cellCount = 0;

  const pushCell = (end: number): boolean => {
    if (row.length >= maxCols) return false;
    row.push(text.slice(cellStart, end));
    cellCount += 1;
    if (cellCount > maxCells) return false;
    return true;
  };

  const pushRow = (): boolean => {
    grid.push(row);
    row = [];
    if (grid.length > maxRows) return false;
    return true;
  };

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    // tab, newline, carriage return
    if (code !== 9 && code !== 10 && code !== 13) continue;

    if (!pushCell(i)) return null;

    if (code === 9) {
      // tab
      cellStart = i + 1;
      continue;
    }

    // newline / CRLF
    if (!pushRow()) return null;

    if (code === 13 && text.charCodeAt(i + 1) === 10) {
      // CRLF -> skip the LF
      i += 1;
      cellStart = i + 1;
    } else {
      cellStart = i + 1;
    }
  }

  // Add the final cell/row (if any). When the payload ends with a newline, `row` will be
  // empty here, mirroring the prior behavior that dropped the final empty record.
  if (cellStart <= text.length) {
    // If we saw at least one delimiter (or any text at all), we have a final cell.
    // Note: for an empty string, this produces `[[""]]` which matches `"".split("\n")`.
    if (row.length > 0 || cellStart < text.length || grid.length === 0) {
      if (!pushCell(text.length)) return null;
      if (!pushRow()) return null;
    }
  }

  // Preserve the previous behavior where a single empty TSV parses as `[[""]]`.
  if (grid.length === 0) return [[""]];
  return grid;
}
