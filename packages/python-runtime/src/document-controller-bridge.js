function rangeSize(range) {
  const rowCount = range.end_row - range.start_row + 1;
  const colCount = range.end_col - range.start_col + 1;
  return { rowCount, colCount };
}

function ensureSingleCell(range) {
  if (range.start_row !== range.end_row || range.start_col !== range.end_col) {
    throw new Error("Expected a single-cell range");
  }
}

/**
 * Adapter that exposes an `apps/desktop` `DocumentController` as the RPC surface
 * expected by the Python `formula` API.
 *
 * This is primarily a convenience bridge; the long-term host integration should
 * wire the Python runtime directly to the core spreadsheet engine.
 */
export class DocumentControllerBridge {
  /**
   * @param {import("../../../apps/desktop/src/document/documentController.js").DocumentController} doc
   * @param {{ activeSheetId?: string }} [options]
   */
  constructor(doc, options = {}) {
    this.doc = doc;
    this.activeSheetId = options.activeSheetId ?? "Sheet1";
    this.sheetIds = new Set([this.activeSheetId]);
    this.selection = {
      sheet_id: this.activeSheetId,
      start_row: 0,
      start_col: 0,
      end_row: 0,
      end_col: 0,
    };
  }

  get_active_sheet_id() {
    return this.activeSheetId;
  }

  get_sheet_id({ name }) {
    return this.sheetIds.has(name) ? name : null;
  }

  create_sheet({ name, index }) {
    const ordered = Array.from(this.sheetIds);

    let insertIndex;
    if (typeof index === "number" && Number.isInteger(index) && index >= 0) {
      insertIndex = Math.min(index, ordered.length);
    } else {
      const activeIdx = ordered.indexOf(this.activeSheetId);
      insertIndex = activeIdx >= 0 ? activeIdx + 1 : ordered.length;
    }

    ordered.splice(insertIndex, 0, name);
    this.sheetIds = new Set(ordered);
    return name;
  }

  get_sheet_name({ sheet_id }) {
    return sheet_id;
  }

  rename_sheet({ sheet_id, name }) {
    if (!this.sheetIds.has(sheet_id)) return null;
    this.sheetIds.delete(sheet_id);
    this.sheetIds.add(name);
    if (this.activeSheetId === sheet_id) this.activeSheetId = name;
    if (this.selection.sheet_id === sheet_id) this.selection.sheet_id = name;
    return null;
  }

  get_selection() {
    return { ...this.selection };
  }

  set_selection({ selection }) {
    if (!selection || !selection.sheet_id) {
      throw new Error("set_selection expects { selection: { sheet_id, start_row, start_col, end_row, end_col } }");
    }
    this.sheetIds.add(selection.sheet_id);
    this.activeSheetId = selection.sheet_id;
    this.selection = { ...selection };
    return null;
  }

  get_range_values({ range }) {
    const values = [];
    for (let r = range.start_row; r <= range.end_row; r++) {
      const rowVals = [];
      for (let c = range.start_col; c <= range.end_col; c++) {
        rowVals.push(this.doc.getCell(range.sheet_id, { row: r, col: c }).value ?? null);
      }
      values.push(rowVals);
    }
    return values;
  }

  set_cell_value({ range, value }) {
    ensureSingleCell(range);
    this.doc.setCellValue(range.sheet_id, { row: range.start_row, col: range.start_col }, value);
    return null;
  }

  get_cell_formula({ range }) {
    ensureSingleCell(range);
    const cell = this.doc.getCell(range.sheet_id, { row: range.start_row, col: range.start_col });
    return cell.formula ?? null;
  }

  set_cell_formula({ range, formula }) {
    ensureSingleCell(range);
    this.doc.setCellFormula(range.sheet_id, { row: range.start_row, col: range.start_col }, formula);
    return null;
  }

  set_range_values({ range, values }) {
    const { rowCount, colCount } = rangeSize(range);
    let matrix;

    if (Array.isArray(values) && Array.isArray(values[0])) {
      matrix = values;
    } else {
      matrix = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => values ?? null));
    }

    // DocumentController expects a CellRange shape: { start: {row,col}, end: {row,col} }
    this.doc.setRangeValues(
      range.sheet_id,
      {
        start: { row: range.start_row, col: range.start_col },
        end: { row: range.end_row, col: range.end_col },
      },
      matrix,
    );
    return null;
  }

  clear_range({ range }) {
    this.doc.clearRange(range.sheet_id, {
      start: { row: range.start_row, col: range.start_col },
      end: { row: range.end_row, col: range.end_col },
    });
    return null;
  }

  set_range_format({ range, format }) {
    const ok = this.doc.setRangeFormat(
      range.sheet_id,
      {
        start: { row: range.start_row, col: range.start_col },
        end: { row: range.end_row, col: range.end_col },
      },
      format,
    );
    if (ok === false) {
      throw new Error("Formatting could not be applied to the full selection. Try selecting fewer cells/rows.");
    }
    return null;
  }

  get_range_format({ range }) {
    const coord = { row: range.start_row, col: range.start_col };

    // Prefer the newer effective (layered) formatting API when available. With layered
    // formatting, `cell.styleId` can be 0 even when an effective format is inherited from
    // sheet/row/column defaults.
    if (typeof this.doc.getCellFormat === "function") {
      const format = this.doc.getCellFormat(range.sheet_id, coord);
      if (format == null) return {};
      // Some controller implementations may return a styleId (number) instead of a style object.
      if (typeof format === "number") {
        return this.doc.styleTable?.get(format) ?? {};
      }
      return format;
    }

    // Back-compat: fall back to the legacy per-cell styleId lookup.
    const cell = this.doc.getCell(range.sheet_id, coord);
    const styleId = cell?.styleId ?? 0;
    return this.doc.styleTable?.get(styleId) ?? {};
  }
}
