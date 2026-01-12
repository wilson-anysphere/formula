import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";
import { LruCache } from "@formula/grid";
import type { DocumentController } from "../../document/documentController.js";

type RichTextValue = { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> };

function isRichTextValue(value: unknown): value is RichTextValue {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function toColumnName(col0: number): string {
  let value = col0 + 1;
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export class DocumentCellProvider implements CellProvider {
  private readonly headerStyle: CellStyle = { fontWeight: "600", textAlign: "center" };
  private readonly rowHeaderStyle: CellStyle = { fontWeight: "600", textAlign: "end" };

  private readonly cache: LruCache<string, CellData | null>;
  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private unsubscribeDoc: (() => void) | null = null;

  constructor(
    private readonly options: {
      document: DocumentController;
      /**
       * Active sheet id for the grid view.
       *
       * The provider is only ever asked for the currently-rendered sheet; callers
       * should update this when switching sheets.
       */
      getSheetId: () => string;
      headerRows: number;
      headerCols: number;
      rowCount: number;
      colCount: number;
      showFormulas: () => boolean;
      getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
      getCommentMeta?: (cellRef: string) => { resolved: boolean } | null;
    }
  ) {
    // Cache covers cell metadata + value formatting work. Keep it bounded to avoid
    // memory blow-ups on huge scrolls.
    this.cache = new LruCache<string, CellData | null>(50_000);
  }

  invalidateAll(): void {
    this.cache.clear();
    for (const listener of this.listeners) listener({ type: "invalidateAll" });
  }

  invalidateDocCells(range: { startRow: number; endRow: number; startCol: number; endCol: number }): void {
    const { headerRows, headerCols } = this.options;
    const gridRange: CellRange = {
      startRow: range.startRow + headerRows,
      endRow: range.endRow + headerRows,
      startCol: range.startCol + headerCols,
      endCol: range.endCol + headerCols
    };

    // Best-effort cache eviction for the affected region.
    const cellCount = Math.max(0, gridRange.endRow - gridRange.startRow) * Math.max(0, gridRange.endCol - gridRange.startCol);
    if (cellCount <= 1000) {
      const sheetId = this.options.getSheetId();
      for (let r = gridRange.startRow; r < gridRange.endRow; r++) {
        for (let c = gridRange.startCol; c < gridRange.endCol; c++) {
          this.cache.delete(`${sheetId}:${r},${c}`);
        }
      }
    } else {
      // If the range is large, clear the whole cache to avoid spending too much time
      // iterating keys.
      this.cache.clear();
    }

    for (const listener of this.listeners) listener({ type: "cells", range: gridRange });
  }

  prefetch(range: CellRange): void {
    // Prefetch is a hint used by async providers. We use it to warm our in-memory cache
    // but cap work so fast scrolls don't block the UI thread.
    const maxCells = 2_000;
    const rows = Math.max(0, range.endRow - range.startRow);
    const cols = Math.max(0, range.endCol - range.startCol);
    const total = rows * cols;
    if (total === 0) return;

    const budget = Math.max(0, Math.min(maxCells, total));
    const step = total / budget;

    let idx = 0;
    while (idx < total) {
      const i = Math.floor(idx);
      const r = range.startRow + Math.floor(i / cols);
      const c = range.startCol + (i % cols);
      this.getCell(r, c);
      idx += step;
    }
  }

  getCell(row: number, col: number): CellData | null {
    const { rowCount, colCount, headerRows, headerCols } = this.options;
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const sheetId = this.options.getSheetId();
    const key = `${sheetId}:${row},${col}`;
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;

    const headerRow = row < headerRows;
    const headerCol = col < headerCols;

    if (headerRow || headerCol) {
      let value: string | number | null = null;
      let style: CellStyle | undefined;
      if (headerRow && headerCol) {
        value = "";
        style = this.headerStyle;
      } else if (headerRow) {
        const docCol = col - headerCols;
        value = docCol >= 0 ? toColumnName(docCol) : "";
        style = this.headerStyle;
      } else {
        const docRow = row - headerRows;
        value = docRow >= 0 ? docRow + 1 : 0;
        style = this.rowHeaderStyle;
      }

      const cell: CellData = { row, col, value, style };
      this.cache.set(key, cell);
      return cell;
    }

    const docRow = row - headerRows;
    const docCol = col - headerCols;

    const state = this.options.document.getCell(sheetId, { row: docRow, col: docCol }) as { value: unknown; formula: string | null };
    if (!state) {
      this.cache.set(key, null);
      return null;
    }

    let value: string | number | boolean | null = null;
    if (state.formula != null) {
      if (this.options.showFormulas()) {
        value = state.formula;
      } else {
        value = this.options.getComputedValue({ row: docRow, col: docCol });
      }
    } else if (state.value != null) {
      value = isRichTextValue(state.value) ? state.value.text : (state.value as any);
      if (value !== null && typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") {
        value = String(state.value);
      }
    }

    const comment = (() => {
      const metaProvider = this.options.getCommentMeta;
      if (!metaProvider) return null;
      const cellRef = `${toColumnName(docCol)}${docRow + 1}`;
      const meta = metaProvider(cellRef);
      if (!meta) return null;
      return { resolved: meta.resolved };
    })();

    const cell: CellData = { row, col, value, comment };
    this.cache.set(key, cell);
    return cell;
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);

    if (!this.unsubscribeDoc) {
      // Coalesce document mutations into provider updates so the renderer can redraw
      // minimal dirty regions.
      this.unsubscribeDoc = this.options.document.on("change", (payload: any) => {
        const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
        if (deltas.length === 0) {
          // Sheet view deltas (frozen panes, row/col sizes, etc.) do not affect cell contents.
          // Avoid evicting the provider cache in those cases; the renderer will be updated by
          // the view sync code (e.g. `syncFrozenPanes` / shared grid axis sync).
          const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
          if (sheetViewDeltas.length > 0 && payload?.recalc !== true) {
            return;
          }
          this.invalidateAll();
          return;
        }

        const sheetId = this.options.getSheetId();
        let minRow = Infinity;
        let maxRow = -Infinity;
        let minCol = Infinity;
        let maxCol = -Infinity;
        let saw = false;

        for (const delta of deltas) {
          if (!delta) continue;
          if (String(delta.sheetId ?? "") !== sheetId) continue;
          const row = Number(delta.row);
          const col = Number(delta.col);
          if (!Number.isInteger(row) || row < 0) continue;
          if (!Number.isInteger(col) || col < 0) continue;
          saw = true;
          minRow = Math.min(minRow, row);
          maxRow = Math.max(maxRow, row);
          minCol = Math.min(minCol, col);
          maxCol = Math.max(maxCol, col);
        }

        if (!saw) {
          this.invalidateAll();
          return;
        }

        this.invalidateDocCells({
          startRow: minRow,
          endRow: maxRow + 1,
          startCol: minCol,
          endCol: maxCol + 1
        });
      });
    }

    return () => {
      this.listeners.delete(listener);
      if (this.listeners.size === 0 && this.unsubscribeDoc) {
        this.unsubscribeDoc();
        this.unsubscribeDoc = null;
      }
    };
  }
}
