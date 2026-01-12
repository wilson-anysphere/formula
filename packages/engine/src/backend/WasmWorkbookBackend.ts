import type { EngineClient } from "../client.ts";
import type { CellScalar } from "../protocol.ts";
import { fromA1, toA1, toA1Range } from "./a1.ts";
import { isFormulaInput, normalizeFormulaTextOpt } from "./formula.ts";
import type { RangeCellEdit, RangeData, SheetInfo, SheetUsedRange, WorkbookBackend, WorkbookInfo } from "@formula/workbook-backend";

type UsedRangeState = {
  start_row: number;
  end_row: number;
  start_col: number;
  end_col: number;
};

const DEFAULT_SHEET: SheetInfo = { id: "Sheet1", name: "Sheet1" };

function isRichTextValue(value: unknown): value is { text: string } {
  return Boolean(value && typeof value === "object" && "text" in value && typeof (value as { text?: unknown }).text === "string");
}

function coerceToScalar(value: unknown): CellScalar {
  if (value == null) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (isRichTextValue(value)) return value.text;
  return null;
}

function cellEditToEngineScalar(edit: RangeCellEdit): CellScalar {
  const formula = typeof edit.formula === "string" ? normalizeFormulaTextOpt(edit.formula) : null;
  if (formula != null) return formula;
  return coerceToScalar(edit.value);
}

function updateUsedRange(map: Map<string, UsedRangeState>, sheetId: string, row: number, col: number): void {
  const existing = map.get(sheetId);
  if (!existing) {
    map.set(sheetId, { start_row: row, end_row: row, start_col: col, end_col: col });
    return;
  }
  existing.start_row = Math.min(existing.start_row, row);
  existing.end_row = Math.max(existing.end_row, row);
  existing.start_col = Math.min(existing.start_col, col);
  existing.end_col = Math.max(existing.end_col, col);
}

type EngineWorkbookJson = {
  sheets?: Record<string, { cells?: Record<string, unknown>; rowCount?: number; colCount?: number }>;
};

export class WasmWorkbookBackend implements WorkbookBackend {
  private readonly usedRanges = new Map<string, UsedRangeState>();
  private workbookInfo: WorkbookInfo | null = null;
  private readonly engine: EngineClient;

  constructor(engine: EngineClient) {
    this.engine = engine;
  }

  async newWorkbook(): Promise<WorkbookInfo> {
    await this.engine.newWorkbook();
    this.usedRanges.clear();

    const info: WorkbookInfo = {
      path: null,
      origin_path: null,
      sheets: [DEFAULT_SHEET],
    };
    this.workbookInfo = info;
    return info;
  }

  async openWorkbookFromBytes(bytes: Uint8Array): Promise<WorkbookInfo> {
    // `loadWorkbookFromXlsxBytes` may transfer/detach the underlying buffer.
    await this.engine.loadWorkbookFromXlsxBytes(bytes);
    await this.engine.recalculate();

    const json = await this.engine.toJson();
    let parsed: EngineWorkbookJson | null = null;
    try {
      parsed = JSON.parse(json) as EngineWorkbookJson;
    } catch {
      parsed = null;
    }

    this.usedRanges.clear();

    const sheetIds = parsed?.sheets && typeof parsed.sheets === "object" ? Object.keys(parsed.sheets) : [];
    for (const sheetId of sheetIds) {
      const cells = parsed?.sheets?.[sheetId]?.cells;
      if (!cells || typeof cells !== "object") continue;

      for (const address of Object.keys(cells)) {
        try {
          const { row0, col0 } = fromA1(address);
          updateUsedRange(this.usedRanges, sheetId, row0, col0);
        } catch {
          // Ignore invalid A1 keys; used range tracking is best-effort for imported workbooks.
        }
      }
    }

    const sheets: SheetInfo[] =
      sheetIds.length > 0 ? sheetIds.map((id) => ({ id, name: id })) : [DEFAULT_SHEET];

    const info: WorkbookInfo = {
      path: null,
      origin_path: null,
      sheets,
    };
    this.workbookInfo = info;
    return info;
  }

  async getSheetUsedRange(sheetId: string): Promise<SheetUsedRange | null> {
    const known = this.usedRanges.get(sheetId);
    return known ? { ...known } : null;
  }

  async getRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  }): Promise<RangeData> {
    const range = toA1Range(params.startRow, params.startCol, params.endRow, params.endCol);
    const result = await this.engine.getRange(range, params.sheetId);

    const values = result.map((row) =>
      row.map((cell) => {
        const input = cell?.input ?? null;
        const formula = isFormulaInput(input) ? normalizeFormulaTextOpt(input) : null;
        const value = cell?.value ?? null;
        return { value, formula, display_value: String(value ?? "") };
      }),
    );

    return { values, start_row: params.startRow, start_col: params.startCol };
  }

  async setCell(params: { sheetId: string; row: number; col: number; value: unknown | null; formula: string | null }): Promise<void> {
    const address = toA1(params.row, params.col);
    const editScalar = cellEditToEngineScalar({ value: params.value, formula: params.formula });
    await this.engine.setCell(address, editScalar, params.sheetId);
    await this.engine.recalculate(params.sheetId);

    if (editScalar != null) {
      updateUsedRange(this.usedRanges, params.sheetId, params.row, params.col);
    }
  }

  async setRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    values: RangeCellEdit[][];
  }): Promise<void> {
    const expectedRows = params.endRow - params.startRow + 1;
    const expectedCols = params.endCol - params.startCol + 1;
    if (params.values.length !== expectedRows || params.values.some((row) => row.length !== expectedCols)) {
      throw new Error(
        `setRange expected ${expectedRows}x${expectedCols} values (got ${params.values.length}x${params.values[0]?.length ?? 0})`,
      );
    }

    const range = toA1Range(params.startRow, params.startCol, params.endRow, params.endCol);
    const scalarValues = params.values.map((row, r) =>
      row.map((cell, c) => {
        const scalar = cellEditToEngineScalar(cell);
        if (scalar != null) {
          updateUsedRange(this.usedRanges, params.sheetId, params.startRow + r, params.startCol + c);
        }
        return scalar;
      }),
    );

    await this.engine.setRange(range, scalarValues, params.sheetId);
    await this.engine.recalculate(params.sheetId);
  }

  // Useful for embedding scenarios (e.g. switching between backends) where callers
  // want to read the last-known workbook metadata.
  getWorkbookInfo(): WorkbookInfo | null {
    return this.workbookInfo;
  }
}
