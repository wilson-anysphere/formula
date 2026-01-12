function rangeSize(range) {
  const startRow = Number(range?.start_row);
  const endRow = Number(range?.end_row);
  const startCol = Number(range?.start_col);
  const endCol = Number(range?.end_col);
  const rowCount = Math.max(0, endRow - startRow + 1);
  const colCount = Math.max(0, endCol - startCol + 1);
  return { rowCount, colCount };
}

function ensureSingleCell(range) {
  if (range.start_row !== range.end_row || range.start_col !== range.end_col) {
    throw new Error("Expected a single-cell range");
  }
}

// Python range APIs frequently return or accept full `Any[][]` matrices. With Excel-scale sheets
// this can easily allocate millions of JS objects and crash the host. Keep reads/writes bounded
// by default to match other extension/scripting safety caps.
const DEFAULT_PYTHON_RANGE_CELL_LIMIT = 200_000;

function assertRangeWithinLimit(range, action) {
  const { rowCount, colCount } = rangeSize(range);
  const cellCount = rowCount * colCount;
  if (!Number.isFinite(cellCount) || cellCount < 0) {
    throw new Error(`${action} skipped: range size is invalid (rows=${rowCount}, cols=${colCount}).`);
  }
  if (cellCount > DEFAULT_PYTHON_RANGE_CELL_LIMIT) {
    throw new Error(
      `${action} skipped for range ${range.sheet_id} (${rowCount}x${colCount}=${cellCount} cells). ` +
        `Limit is ${DEFAULT_PYTHON_RANGE_CELL_LIMIT} cells.`
    );
  }
}

function normalizeSheetNameForCaseInsensitiveCompare(name) {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics used by the desktop sheet store / workbook backend.
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
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
    const initialSheetIds = typeof doc?.getSheetIds === "function" ? doc.getSheetIds() : [];
    this.activeSheetId = options.activeSheetId ?? initialSheetIds[0] ?? "Sheet1";
    this.sheetIds = new Set(initialSheetIds.length > 0 ? initialSheetIds : [this.activeSheetId]);
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
    const desired = normalizeSheetNameForCaseInsensitiveCompare(name);
    const docIds = typeof this.doc?.getSheetIds === "function" ? this.doc.getSheetIds() : null;
    const ids = docIds && docIds.length > 0 ? docIds : Array.from(this.sheetIds);

    // Back-compat: some callers may still pass a sheet id directly.
    if (ids.includes(name)) return name;

    for (const sheetId of ids) {
      const meta = typeof this.doc?.getSheetMeta === "function" ? this.doc.getSheetMeta(sheetId) : null;
      const sheetName = meta?.name ?? sheetId;
      if (normalizeSheetNameForCaseInsensitiveCompare(sheetName) === desired) return sheetId;
    }
    return null;
  }

  create_sheet({ name, index }) {
    if (typeof this.doc?.addSheet !== "function") {
      // Legacy fallback: treat sheet id and name as the same string.
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

    let beforeOrder = typeof this.doc?.getSheetIds === "function" ? this.doc.getSheetIds() : Array.from(this.sheetIds);

    // DocumentController materializes sheets lazily. If the host hasn't touched the active sheet yet
    // (so `getSheetIds()` is empty) but our bridge has an `activeSheetId`, ensure the active sheet
    // exists before inserting a new sheet "after active". This matches user expectations from Excel,
    // where the active sheet always exists even if no cells have been accessed.
    if (
      beforeOrder.length === 0 &&
      typeof this.doc?.getSheetIds === "function" &&
      typeof this.doc?.getCell === "function" &&
      this.activeSheetId
    ) {
      try {
        this.doc.getCell(this.activeSheetId, { row: 0, col: 0 });
        const refreshed = this.doc.getSheetIds();
        if (Array.isArray(refreshed) && refreshed.length > 0) {
          beforeOrder = refreshed;
          this.sheetIds = new Set(refreshed);
        }
      } catch {
        // ignore
      }
    }

    const hasExplicitIndex = typeof index === "number" && Number.isInteger(index) && index >= 0;

    if (!hasExplicitIndex) {
      const insertAfterId = beforeOrder.includes(this.activeSheetId) ? this.activeSheetId : null;
      const newId = this.doc.addSheet({ name, insertAfterId });
      this.sheetIds = new Set(this.doc.getSheetIds?.() ?? [...this.sheetIds, newId]);
      return newId;
    }

    const clampedIndex = Math.min(index, beforeOrder.length);

    // DocumentController only supports "insert after" semantics. For index=0, we create
    // the sheet and then reorder to the desired absolute position.
    if (clampedIndex === 0 && beforeOrder.length > 0 && typeof this.doc?.reorderSheets === "function") {
      const canBatch = typeof this.doc?.beginBatch === "function" && typeof this.doc?.endBatch === "function";
      if (canBatch) this.doc.beginBatch({ label: "Add Sheet" });
      try {
        const newId = this.doc.addSheet({ name, insertAfterId: null });
        this.doc.reorderSheets([newId, ...beforeOrder]);
        if (canBatch) this.doc.endBatch();
        this.sheetIds = new Set(this.doc.getSheetIds?.() ?? [...this.sheetIds, newId]);
        return newId;
      } catch (err) {
        if (canBatch && typeof this.doc?.cancelBatch === "function") {
          try {
            this.doc.cancelBatch();
          } catch {
            // ignore
          }
        }
        throw err;
      }
    }

    const insertAfterId = clampedIndex > 0 ? beforeOrder[clampedIndex - 1] ?? null : null;
    const newId = this.doc.addSheet({ name, insertAfterId });
    this.sheetIds = new Set(this.doc.getSheetIds?.() ?? [...this.sheetIds, newId]);
    return newId;
  }

  get_sheet_name({ sheet_id }) {
    const meta = typeof this.doc?.getSheetMeta === "function" ? this.doc.getSheetMeta(sheet_id) : null;
    return meta?.name ?? sheet_id;
  }

  rename_sheet({ sheet_id, name }) {
    if (typeof this.doc?.renameSheet === "function") {
      this.doc.renameSheet(sheet_id, name);
      return null;
    }

    // Legacy fallback where sheet id and name are treated as the same string.
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
    assertRangeWithinLimit(range, "get_range_values");
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
    assertRangeWithinLimit(range, "set_range_values");
    const { rowCount, colCount } = rangeSize(range);
    const isSingleCellRange = rowCount === 1 && colCount === 1;
    let matrix;

    if (Array.isArray(values) && Array.isArray(values[0])) {
      // Validate that the provided 2D array fits inside the declared range. We intentionally
      // allow short rows (missing values become null), but do not allow writing beyond the
      // range bounds (which could trigger large allocations by accident).
      const providedRowCount = values.length;
      let providedColCount = 0;
      for (const row of values) {
        if (!Array.isArray(row)) continue;
        if (row.length > providedColCount) providedColCount = row.length;
      }
      const cellCount = providedRowCount * providedColCount;
      if (cellCount > DEFAULT_PYTHON_RANGE_CELL_LIMIT) {
        throw new Error(
          `set_range_values skipped for values matrix (${providedRowCount}x${providedColCount}=${cellCount} cells). ` +
            `Limit is ${DEFAULT_PYTHON_RANGE_CELL_LIMIT} cells.`
        );
      }
      // Spill behavior: when the destination is a single cell, treat the range as a start
      // anchor (matching Excel/VBA semantics) and allow the values matrix to expand the
      // written rectangle.
      if (!isSingleCellRange && (providedRowCount > rowCount || providedColCount > colCount)) {
        throw new Error(
          `set_range_values values shape (${providedRowCount}x${providedColCount}) exceeds range shape (${rowCount}x${colCount}). ` +
            `Select a smaller values matrix or a larger destination range.`
        );
      }
      matrix = values;
    } else {
      // Scalar fill. Use a shared row buffer to avoid allocating rowCount*colCount JS values.
      const row = Array.from({ length: colCount }, () => values ?? null);
      matrix = Array.from({ length: rowCount }, () => row);
    }

    // DocumentController accepts either:
    // - a CellCoord start cell (range inferred from values dimensions), or
    // - a CellRange (explicit rectangle).
    //
    // Use start-cell semantics for single-cell ranges so 2D assignments "spill" from the anchor cell.
    this.doc.setRangeValues(
      range.sheet_id,
      isSingleCellRange && Array.isArray(matrix) && Array.isArray(matrix[0])
        ? { row: range.start_row, col: range.start_col }
        : {
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
