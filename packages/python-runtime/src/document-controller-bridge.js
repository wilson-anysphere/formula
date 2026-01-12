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

    const beforeOrder = typeof this.doc?.getSheetIds === "function" ? this.doc.getSheetIds() : Array.from(this.sheetIds);

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
