import { CellEditorOverlay } from "../editor/cellEditorOverlay";
import { cellToA1, rangeToA1 } from "../selection/a1";
import { navigateSelectionByKey } from "../selection/navigation";
import { SelectionRenderer } from "../selection/renderer";
import type { CellCoord, GridLimits, Range, SelectionState } from "../selection/types";
import {
  DEFAULT_GRID_LIMITS,
  addCellToSelection,
  createSelection,
  extendSelectionToCell,
  selectAll,
  selectColumns,
  selectRows,
  setActiveCell
} from "../selection/selection";
import { DocumentController } from "../document/documentController.js";
import { MockEngine } from "../document/engine.js";

export interface SpreadsheetAppStatusElements {
  activeCell: HTMLElement;
  selectionRange: HTMLElement;
  activeValue: HTMLElement;
}

export class SpreadsheetApp {
  private readonly sheetId = "Sheet1";
  private readonly engine = new MockEngine();
  private readonly document = new DocumentController({ engine: this.engine });
  private limits: GridLimits;

  private gridCanvas: HTMLCanvasElement;
  private selectionCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private selectionCtx: CanvasRenderingContext2D;

  private dpr = 1;
  private width = 0;
  private height = 0;

  private readonly cellWidth = 100;
  private readonly cellHeight = 24;

  private selection: SelectionState;
  private selectionRenderer = new SelectionRenderer();

  private editor: CellEditorOverlay;

  private resizeObserver: ResizeObserver;

  constructor(private root: HTMLElement, private status: SpreadsheetAppStatusElements, opts: { limits?: GridLimits } = {}) {
    this.limits = opts.limits ?? { ...DEFAULT_GRID_LIMITS, maxRows: 10_000, maxCols: 200 };
    this.selection = createSelection({ row: 0, col: 0 }, this.limits);

    // Seed data for navigation tests (used range ends at D5).
    this.document.setCellValue(this.sheetId, { row: 0, col: 0 }, "Seed");
    this.document.setCellValue(this.sheetId, { row: 4, col: 3 }, "BottomRight");

    this.gridCanvas = document.createElement("canvas");
    this.gridCanvas.className = "grid-canvas";
    this.gridCanvas.setAttribute("aria-hidden", "true");
    this.selectionCanvas = document.createElement("canvas");
    this.selectionCanvas.className = "grid-canvas";
    this.selectionCanvas.setAttribute("aria-hidden", "true");

    this.root.appendChild(this.gridCanvas);
    this.root.appendChild(this.selectionCanvas);

    const gridCtx = this.gridCanvas.getContext("2d");
    const selectionCtx = this.selectionCanvas.getContext("2d");
    if (!gridCtx || !selectionCtx) {
      throw new Error("Canvas 2D context not available");
    }
    this.gridCtx = gridCtx;
    this.selectionCtx = selectionCtx;

    this.editor = new CellEditorOverlay(this.root, {
      onCommit: (commit) => {
        this.applyEdit(commit.cell, commit.value);
        this.renderGrid();

        const next = navigateSelectionByKey(
          this.selection,
          commit.reason === "enter" ? "Enter" : "Tab",
          { shift: commit.shift, primary: false },
          this.usedRangeProvider(),
          this.limits
        );

        if (next) this.selection = next;
        this.renderSelection();
        this.updateStatus();
        this.focus();
      },
      onCancel: () => {
        this.renderSelection();
        this.updateStatus();
        this.focus();
      }
    });

    this.root.addEventListener("pointerdown", (e) => this.onPointerDown(e));
    this.root.addEventListener("keydown", (e) => this.onKeyDown(e));

    this.resizeObserver = new ResizeObserver(() => this.onResize());
    this.resizeObserver.observe(this.root);

    // Initial layout + render.
    this.onResize();
  }

  destroy(): void {
    this.resizeObserver.disconnect();
    this.root.replaceChildren();
  }

  focus(): void {
    this.root.focus();
  }

  getRecalcCount(): number {
    return this.engine.recalcCount;
  }

  getCellValueA1(a1: string): string {
    const cell = parseA1(a1);
    return this.getCellDisplayValue(cell);
  }

  private onResize(): void {
    const rect = this.root.getBoundingClientRect();
    this.width = rect.width;
    this.height = rect.height;
    this.dpr = window.devicePixelRatio || 1;

    for (const canvas of [this.gridCanvas, this.selectionCanvas]) {
      canvas.width = Math.floor(this.width * this.dpr);
      canvas.height = Math.floor(this.height * this.dpr);
      canvas.style.width = `${this.width}px`;
      canvas.style.height = `${this.height}px`;
    }

    // Reset transforms and apply DPR scaling so drawing code uses CSS pixels.
    for (const ctx of [this.gridCtx, this.selectionCtx]) {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.scale(this.dpr, this.dpr);
    }

    this.renderGrid();
    this.renderSelection();
    this.updateStatus();
  }

  private renderGrid(): void {
    const ctx = this.gridCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.gridCanvas.width, this.gridCanvas.height);
    ctx.restore();

    ctx.save();
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, this.width, this.height);

    const cols = Math.max(1, Math.floor(this.width / this.cellWidth));
    const rows = Math.max(1, Math.floor(this.height / this.cellHeight));

    ctx.strokeStyle = "#d4d4d4";
    ctx.lineWidth = 1;

    for (let r = 0; r <= rows; r++) {
      const y = r * this.cellHeight + 0.5;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(cols * this.cellWidth, y);
      ctx.stroke();
    }

    for (let c = 0; c <= cols; c++) {
      const x = c * this.cellWidth + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, rows * this.cellHeight);
      ctx.stroke();
    }

    ctx.fillStyle = "#000";
    ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const value = this.getCellDisplayValue({ row: r, col: c });
        if (value === "") continue;
        ctx.fillText(value, c * this.cellWidth + 4, r * this.cellHeight + 16);
      }
    }

    ctx.restore();
  }

  private renderSelection(): void {
    this.selectionRenderer.render(this.selectionCtx, this.selection, {
      getCellRect: (cell) => this.getCellRect(cell)
    });

    // If scrolling/resizing happened during editing, keep the editor aligned.
    if (this.editor.isOpen()) {
      this.editor.reposition(this.getCellRect(this.selection.active));
    }
  }

  private updateStatus(): void {
    this.status.activeCell.textContent = cellToA1(this.selection.active);
    this.status.selectionRange.textContent =
      this.selection.ranges.length === 1 ? rangeToA1(this.selection.ranges[0]) : `${this.selection.ranges.length} ranges`;
    this.status.activeValue.textContent = this.getCellDisplayValue(this.selection.active);
  }

  private getCellRect(cell: CellCoord) {
    return {
      x: cell.col * this.cellWidth,
      y: cell.row * this.cellHeight,
      width: this.cellWidth,
      height: this.cellHeight
    };
  }

  private cellFromPoint(pointX: number, pointY: number): CellCoord {
    const col = Math.floor(pointX / this.cellWidth);
    const row = Math.floor(pointY / this.cellHeight);
    return {
      row: Math.max(0, Math.min(this.limits.maxRows - 1, row)),
      col: Math.max(0, Math.min(this.limits.maxCols - 1, col))
    };
  }

  private onPointerDown(e: PointerEvent): void {
    if (this.editor.isOpen()) return;

    const rect = this.root.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const cell = this.cellFromPoint(x, y);

    if (e.shiftKey) {
      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
    } else if (e.ctrlKey || e.metaKey) {
      this.selection = addCellToSelection(this.selection, cell, this.limits);
    } else {
      this.selection = setActiveCell(this.selection, cell, this.limits);
    }

    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (this.editor.isOpen()) {
      // The editor handles Enter/Tab/Escape itself. We keep focus on the textarea.
      return;
    }

    // Editing
    if (e.key === "F2") {
      e.preventDefault();
      const cell = this.selection.active;
      const bounds = this.getCellRect(cell);
      const initialValue = this.getCellDisplayValue(cell);
      this.editor.open(cell, bounds, initialValue, { cursor: "end" });
      return;
    }

    // Selection shortcuts
    const primary = e.ctrlKey || e.metaKey;
    if (primary && (e.key === "a" || e.key === "A")) {
      e.preventDefault();
      this.selection = selectAll(this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (primary && e.code === "Space") {
      // Ctrl+Space selects entire column.
      e.preventDefault();
      this.selection = selectColumns(this.selection, this.selection.active.col, this.selection.active.col, {}, this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    if (!primary && e.shiftKey && e.code === "Space") {
      // Shift+Space selects entire row.
      e.preventDefault();
      this.selection = selectRows(this.selection, this.selection.active.row, this.selection.active.row, {}, this.limits);
      this.renderSelection();
      this.updateStatus();
      return;
    }

    const next = navigateSelectionByKey(
      this.selection,
      e.key,
      { shift: e.shiftKey, primary },
      this.usedRangeProvider(),
      this.limits
    );
    if (!next) return;

    e.preventDefault();
    this.selection = next;
    this.renderSelection();
    this.updateStatus();
  }

  private getCellDisplayValue(cell: CellCoord): string {
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (state?.value != null) return String(state.value);
    if (state?.formula != null) return `=${state.formula}`;
    return "";
  }

  private applyEdit(cell: CellCoord, rawValue: string): void {
    if (rawValue.trim() === "") {
      this.document.clearCell(this.sheetId, cell, { label: "Clear cell" });
      return;
    }

    if (rawValue.startsWith("=")) {
      this.document.setCellFormula(this.sheetId, cell, rawValue.slice(1), { label: "Edit cell" });
      return;
    }

    this.document.setCellValue(this.sheetId, cell, rawValue, { label: "Edit cell" });
  }

  private usedRangeProvider() {
    return {
      getUsedRange: () => this.computeUsedRange(),
      isCellEmpty: (cell: CellCoord) => {
        const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
        return state?.value == null && state?.formula == null;
      }
    };
  }

  private computeUsedRange(): Range | null {
    const sheet = (this.document as any).model?.sheets?.get(this.sheetId) as { cells: Map<string, any> } | undefined;
    if (!sheet) return null;
    if (!sheet.cells || sheet.cells.size === 0) return null;

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;
    let hasData = false;

    for (const [key, cell] of sheet.cells.entries()) {
      if (cell?.value == null && cell?.formula == null) continue;
      const [rowStr, colStr] = String(key).split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
      hasData = true;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }

    if (!hasData) return null;
    return { startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol };
  }
}

function parseA1(a1: string): CellCoord {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(a1.trim());
  if (!match) return { row: 0, col: 0 };
  const colName = match[1].toUpperCase();
  const row = Number(match[2]) - 1;
  let col = 0;
  for (let i = 0; i < colName.length; i++) {
    col = col * 26 + (colName.charCodeAt(i) - 64);
  }
  return { row: Math.max(0, row), col: Math.max(0, col - 1) };
}
