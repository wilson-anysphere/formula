import { normalizeName } from "../../../../packages/search/index.js";
import { parseImageCellValue } from "../shared/imageCellValue.js";

function parseCellKey(key) {
  const [r, c] = String(key).split(",");
  const row = Number.parseInt(r, 10);
  const col = Number.parseInt(c, 10);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    return null;
  }
  return { row, col };
}

function formatValueForDisplay(value) {
  if (value == null) return "";

  // DocumentController stores rich text as `{ text, runs }`. Search UI should treat this as
  // plain text instead of rendering `[object Object]`.
  if (typeof value === "object" && typeof value.text === "string") {
    return value.text;
  }

  const image = parseImageCellValue(value);
  if (image) return image.altText ?? "[Image]";

  return String(value);
}

function normalizeFormula(formula) {
  if (formula == null) return null;
  const trimmed = String(formula).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

function normalizeSheetNameForCaseInsensitiveCompare(name) {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics used by the desktop backend + workbook-backend shared validator.
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

/**
 * Adapter that exposes a DocumentController-like model through the interface expected
 * by `packages/search` (workbook -> sheets -> cells).
 *
 * This keeps `packages/search` UI-agnostic while still allowing the desktop app
 * to reuse the same search/replace implementation.
 */
export class DocumentWorkbookAdapter {
  /**
   * @param {{
   *   document: import("../document/documentController.js").DocumentController,
   *   sheetNameResolver?: import("../sheet/sheetNameResolver.ts").SheetNameResolver,
   * }} params
   */
  constructor({ document, sheetNameResolver } = {}) {
    if (!document) throw new Error("DocumentWorkbookAdapter: document is required");
    this.document = document;
    this.sheetNameResolver = sheetNameResolver ?? null;
    /** @type {Map<string, DocumentSheetAdapter>} */
    this.#sheetsById = new Map();

    /**
     * Cache-buster for workbook metadata (defined names / tables).
     *
     * Consumers like tab-completion cache suggestions and need a cheap way to
     * invalidate when metadata changes.
     */
    this.schemaVersion = 0;

    /** @type {Map<string, any>} */
    this.names = new Map();
    /** @type {Map<string, any>} */
    this.tables = new Map();
  }

  /** @type {Map<string, DocumentSheetAdapter>} */
  #sheetsById;

  get sheets() {
    const ids = typeof this.document.getSheetIds === "function" ? this.document.getSheetIds() : [];
    if (ids.length === 0) return [];

    return ids.map((id) => this.#getSheetById(id));
  }

  getSheet(sheetName) {
    const sheetId = this.#resolveSheetIdByName(sheetName);
    if (!sheetId) {
      throw new Error(`Unknown sheet: ${sheetName}`);
    }
    return this.#getSheetById(sheetId);
  }

  defineName(name, ref) {
    this.schemaVersion += 1;
    this.names.set(normalizeName(name), { ...ref, name: String(name) });
  }

  getName(name) {
    return this.names.get(normalizeName(name)) ?? null;
  }

  addTable(table) {
    this.schemaVersion += 1;
    this.tables.set(normalizeName(table.name), table);
  }

  getTable(name) {
    return this.tables.get(normalizeName(name)) ?? null;
  }

  clearSchema() {
    this.schemaVersion += 1;
    this.names.clear();
    this.tables.clear();
  }

  dispose() {
    // Best-effort memory release: clear cached sheet adapters + schema so a workbook adapter
    // doesn't retain large name/table metadata after a SpreadsheetApp teardown.
    try {
      this.clearSchema();
    } catch {
      // ignore
    }
    try {
      this.#sheetsById.clear();
    } catch {
      // ignore
    }
  }

  #resolveSheetIdByName(sheetName) {
    const trimmed = (() => {
      const raw = String(sheetName ?? "").trim();
      // Excel-style sheet-qualified references may be quoted, e.g. `'My Sheet'!A1`.
      // When callers pass the sheet token through directly, normalize it back to
      // the display name before resolving.
      const quoted = /^'((?:[^']|'')+)'$/.exec(raw);
      if (quoted) return quoted[1].replace(/''/g, "'").trim();
      return raw;
    })();
    if (!trimmed) return null;

    const resolved = this.sheetNameResolver?.getSheetIdByName?.(trimmed);
    if (typeof resolved === "string") {
      const id = resolved.trim();
      if (id) return id;
    }

    // Pass through stable ids when the resolver recognizes them (even if the sheet
    // hasn't been created in the DocumentController yet).
    const display = this.sheetNameResolver?.getSheetNameById?.(trimmed);
    if (typeof display === "string" && display.trim()) return trimmed;

    // Fallback: allow referring to sheets by id (legacy behavior) and, when possible, by the
    // DocumentController's metadata name (useful when callers haven't provided a separate
    // sheetNameResolver but the document still has user-facing names stored in sheet meta).
    const ids = typeof this.document.getSheetIds === "function" ? this.document.getSheetIds() : [];
    const needleIdCi = trimmed.toLowerCase();
    const needleNameCi = normalizeSheetNameForCaseInsensitiveCompare(trimmed);

    const byId = ids.find((id) => String(id).toLowerCase() === needleIdCi) ?? null;
    if (byId) return byId;

    if (typeof this.document.getSheetMeta === "function") {
      for (const id of ids) {
        const meta = this.document.getSheetMeta(id);
        const name = meta?.name;
        if (
          typeof name === "string" &&
          normalizeSheetNameForCaseInsensitiveCompare(name.trim()) === needleNameCi
        ) {
          return id;
        }
      }
    }

    return null;
  }

  #getSheetById(sheetId) {
    const key = String(sheetId);
    let sheet = this.#sheetsById.get(key);
    if (!sheet) {
      sheet = new DocumentSheetAdapter(this.document, key, this.sheetNameResolver);
      this.#sheetsById.set(key, sheet);
    }
    return sheet;
  }
}

class DocumentSheetAdapter {
  /**
   * @param {import("../document/documentController.js").DocumentController} document
   * @param {string} sheetId
   * @param {import("../sheet/sheetNameResolver.ts").SheetNameResolver | null} sheetNameResolver
   */
  constructor(document, sheetId, sheetNameResolver) {
    this.document = document;
    this.sheetId = sheetId;
    this.sheetNameResolver = sheetNameResolver ?? null;
  }

  get name() {
    const resolved = this.sheetNameResolver?.getSheetNameById?.(this.sheetId);
    if (resolved) return resolved;
    const metaName = this.document.getSheetMeta?.(this.sheetId)?.name;
    if (typeof metaName === "string" && metaName.trim() !== "") return metaName;
    return this.sheetId;
  }

  getUsedRange() {
    if (typeof this.document.getUsedRange === "function") {
      return this.document.getUsedRange(this.sheetId);
    }
    return null;
  }

  /**
   * Merged-cell regions for this sheet (inclusive coordinates).
   *
   * `packages/search` uses this to expand selection scopes and to avoid double-indexing
   * merged regions (Excel-style: only the top-left anchor cell is searchable).
   */
  getMergedRanges() {
    const doc = /** @type {any} */ (this.document);
    // Avoid creating phantom sheets when callers hold stale sheet adapters.
    if (typeof doc.getSheetMeta === "function") {
      const meta = doc.getSheetMeta(this.sheetId);
      if (!meta) return [];
    } else if (doc?.model?.sheets instanceof Map && !doc.model.sheets.has(this.sheetId)) {
      return [];
    }
    if (typeof doc.getMergedRanges === "function") {
      return doc.getMergedRanges(this.sheetId) ?? [];
    }
    // Fallback: read from sheet view state if the controller does not expose a helper.
    if (typeof doc.getSheetView === "function") {
      const view = doc.getSheetView(this.sheetId);
      const ranges = view?.mergedRanges ?? view?.mergedCells;
      return Array.isArray(ranges) ? ranges : [];
    }
    return [];
  }

  /**
   * Resolve a cell inside a merged region to its anchor (top-left).
   *
   * @param {number} row
   * @param {number} col
   * @returns {{ row: number, col: number } | null}
   */
  getMergedMasterCell(row, col) {
    const doc = /** @type {any} */ (this.document);
    // Avoid creating phantom sheets when callers hold stale sheet adapters.
    if (typeof doc.getSheetMeta === "function") {
      const meta = doc.getSheetMeta(this.sheetId);
      if (!meta) return null;
    } else if (doc?.model?.sheets instanceof Map && !doc.model.sheets.has(this.sheetId)) {
      return null;
    }
    if (typeof doc.getMergedMasterCell === "function") {
      return doc.getMergedMasterCell(this.sheetId, { row, col });
    }
    const ranges = this.getMergedRanges();
    if (!ranges || ranges.length === 0) return null;
    for (const r of ranges) {
      if (!r) continue;
      if (row < r.startRow || row > r.endRow) continue;
      if (col < r.startCol || col > r.endCol) continue;
      return { row: r.startRow, col: r.startCol };
    }
    return null;
  }

  getCell(row, col) {
    const doc = /** @type {any} */ (this.document);
    const coord = { row, col };
    // `DocumentController.getCell()` creates sheets lazily. Prefer `peekCell` when available so
    // probing for values does not resurrect deleted sheets or create phantom sheets from a
    // stale sheetNameResolver mapping.
    const state =
      typeof doc.peekCell === "function"
        ? doc.peekCell(this.sheetId, coord)
        : doc?.model?.sheets instanceof Map && !doc.model.sheets.has(this.sheetId)
          ? null
          : doc.getCell(this.sheetId, coord);
    if (!state) return null;
    const formula = normalizeFormula(state.formula);
    const value = state.value ?? null;

    if (value == null && formula == null) return null;

    const display = value != null ? formatValueForDisplay(value) : formula ?? "";
    return { value, formula, display };
  }

  setCell(row, col, cell) {
    const doc = /** @type {any} */ (this.document);
    // Avoid resurrecting deleted sheets (or creating new ones) via search/replace edits.
    if (typeof doc.getSheetMeta === "function") {
      const meta = doc.getSheetMeta(this.sheetId);
      if (!meta) {
        throw new Error(`Unknown sheet: ${this.sheetId}`);
      }
    } else if (doc?.model?.sheets instanceof Map && !doc.model.sheets.has(this.sheetId)) {
      throw new Error(`Unknown sheet: ${this.sheetId}`);
    }

    if (!cell || (cell.value == null && (cell.formula == null || cell.formula === ""))) {
      this.document.clearCell(this.sheetId, { row, col });
      return;
    }

      if (cell.formula != null && cell.formula !== "") {
        const formula = String(cell.formula);
        const normalized = normalizeFormula(formula);
        // DocumentController stores formulas with a leading "=". Normalize here so callers
        // can provide formula text with or without "=" (e.g. from user edits/search replace).
        this.document.setCellFormula(this.sheetId, { row, col }, normalized);
        return;
      }

    this.document.setCellValue(this.sheetId, { row, col }, cell.value);
  }

  *iterateCells(range, { order = "byRows" } = {}) {
    const used = this.getUsedRange();
    if (!used) return;

    const startRow = Math.max(used.startRow, range.startRow);
    const endRow = Math.min(used.endRow, range.endRow);
    const startCol = Math.max(used.startCol, range.startCol);
    const endCol = Math.min(used.endCol, range.endCol);
    if (startRow > endRow || startCol > endCol) return;

    const results = [];
    if (typeof this.document.forEachCellInSheet === "function") {
      // Use the DocumentController's cell iterator to avoid re-parsing sparse map keys.
      this.document.forEachCellInSheet(this.sheetId, ({ row, col, cell }) => {
        if (!cell) return;
        if (cell.value == null && cell.formula == null) return;
        if (row < startRow || row > endRow || col < startCol || col > endCol) return;
        const formula = normalizeFormula(cell.formula);
        const value = cell.value ?? null;
        const display = value != null ? formatValueForDisplay(value) : formula ?? "";
        results.push({ row, col, cell: { value, formula, display } });
      });
    } else {
      // Fallback path for older document implementations.
      const sheet = this.document.model?.sheets?.get(this.sheetId);
      for (const [key, cell] of sheet?.cells?.entries?.() ?? []) {
        if (!cell) continue;
        if (cell.value == null && cell.formula == null) continue;
        const parsed = parseCellKey(key);
        if (!parsed) continue;
        if (parsed.row < startRow || parsed.row > endRow || parsed.col < startCol || parsed.col > endCol) continue;

        const formula = normalizeFormula(cell.formula);
        const value = cell.value ?? null;
        const display = value != null ? formatValueForDisplay(value) : formula ?? "";
        results.push({ row: parsed.row, col: parsed.col, cell: { value, formula, display } });
      }
    }

    results.sort((a, b) => {
      if (order === "byColumns") {
        if (a.col !== b.col) return a.col - b.col;
        return a.row - b.row;
      }
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });

    for (const entry of results) yield entry;
  }
}
