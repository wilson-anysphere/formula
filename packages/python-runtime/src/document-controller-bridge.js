import { cloneCellState } from "../../../apps/desktop/src/document/cell.js";

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
  }

  get_active_sheet_id() {
    return this.activeSheetId;
  }

  get_sheet_id({ name }) {
    return this.sheetIds.has(name) ? name : null;
  }

  create_sheet({ name }) {
    this.sheetIds.add(name);
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

  /**
   * Convenience helper for tests/debugging.
   */
  dump_cell_state({ sheet_id, row, col }) {
    return cloneCellState(this.doc.getCell(sheet_id, { row, col }));
  }
}

