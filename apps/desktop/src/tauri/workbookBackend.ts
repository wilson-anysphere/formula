import type {
  DefinedNameInfo,
  RangeCellEdit,
  RangeData,
  SheetUsedRange,
  TableInfo,
  WorkbookBackend,
  WorkbookInfo,
  WorkbookThemePalette,
} from "@formula/workbook-backend";

export type {
  CellValue,
  DefinedNameInfo,
  RangeCellEdit,
  RangeData,
  SheetInfo,
  SheetVisibility,
  SheetUsedRange,
  TabColor,
  TableInfo,
  WorkbookBackend,
  WorkbookInfo,
  WorkbookThemePalette,
} from "@formula/workbook-backend";

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export class TauriWorkbookBackend implements WorkbookBackend {
  private readonly invoke: TauriInvoke;

  constructor() {
    this.invoke = getTauriInvoke();
  }

  async newWorkbook(): Promise<WorkbookInfo> {
    const info = await this.invoke("new_workbook");
    return info as WorkbookInfo;
  }

  async openWorkbook(path: string): Promise<WorkbookInfo> {
    const info = await this.invoke("open_workbook", { path });
    return info as WorkbookInfo;
  }

  async moveSheet(sheetId: string, toIndex: number): Promise<void> {
    await this.invoke("move_sheet", { sheet_id: sheetId, to_index: toIndex });
  }

  async getWorkbookThemePalette(): Promise<WorkbookThemePalette | null> {
    const palette = await this.invoke("get_workbook_theme_palette");
    return (palette as WorkbookThemePalette | null) ?? null;
  }

  async listDefinedNames(): Promise<DefinedNameInfo[]> {
    const payload = await this.invoke("list_defined_names");
    return (payload as DefinedNameInfo[]) ?? [];
  }

  async listTables(): Promise<TableInfo[]> {
    const payload = await this.invoke("list_tables");
    return (payload as TableInfo[]) ?? [];
  }

  /**
   * Desktop-only: fetch chart models parsed from the opened XLSX package.
   *
   * These are JSON-serialized Rust `formula_model::charts::ChartModel` values keyed by a
   * stable `chart_id` (recommended: `${sheetId}:${drawingObjectId}`).
   */
  async listImportedChartModels(): Promise<unknown[]> {
    const payload = await this.invoke("list_imported_chart_models");
    return (payload as unknown[]) ?? [];
  }

  /**
   * Desktop-only: fetch chart drawing objects parsed from the opened XLSX package.
   *
   * Each entry includes the DrawingML anchor (for positioning) and may include a parsed chart
   * `model` (for rendering). The frontend uses this to populate both the drawings overlay and the
   * imported chart model store.
   */
  async listImportedChartObjects(): Promise<unknown[]> {
    const payload = await this.invoke("list_imported_chart_objects");
    return (payload as unknown[]) ?? [];
  }

  async saveWorkbook(path?: string): Promise<void> {
    const args: Record<string, unknown> = {};
    if (path) args.path = path;
    await this.invoke("save_workbook", args);
  }

  async getRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  }): Promise<RangeData> {
    const payload = await this.invoke("get_range", {
      sheet_id: params.sheetId,
      start_row: params.startRow,
      start_col: params.startCol,
      end_row: params.endRow,
      end_col: params.endCol
    });
    return payload as RangeData;
  }

  async getSheetUsedRange(sheetId: string): Promise<SheetUsedRange | null> {
    const payload = await this.invoke("get_sheet_used_range", {
      sheet_id: sheetId,
    });
    return (payload as SheetUsedRange | null) ?? null;
  }

  /**
   * Best-effort: not all backend builds expose formatting snapshot commands.
   * Callers should tolerate failures and treat them as "no persisted formatting".
   */
  async getSheetFormatting(sheetId: string): Promise<unknown | null> {
    const payload = await this.invoke("get_sheet_formatting", {
      sheet_id: sheetId,
    });
    return (payload as unknown | null) ?? null;
  }

  async setCell(params: {
    sheetId: string;
    row: number;
    col: number;
    value: unknown | null;
    formula: string | null;
  }): Promise<void> {
    await this.invoke("set_cell", {
      sheet_id: params.sheetId,
      row: params.row,
      col: params.col,
      value: params.value,
      formula: params.formula
    });
  }

  async setRange(params: {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    values: RangeCellEdit[][];
  }): Promise<void> {
    await this.invoke("set_range", {
      sheet_id: params.sheetId,
      start_row: params.startRow,
      start_col: params.startCol,
      end_row: params.endRow,
      end_col: params.endCol,
      values: params.values
    });
  }

  async getPrecedents(params: {
    sheetId: string;
    row: number;
    col: number;
    transitive?: boolean;
  }): Promise<string[]> {
    const payload = await this.invoke("get_precedents", {
      sheet_id: params.sheetId,
      row: params.row,
      col: params.col,
      transitive: params.transitive ?? false,
    });
    return (payload as string[]) ?? [];
  }

  async getDependents(params: {
    sheetId: string;
    row: number;
    col: number;
    transitive?: boolean;
  }): Promise<string[]> {
    const payload = await this.invoke("get_dependents", {
      sheet_id: params.sheetId,
      row: params.row,
      col: params.col,
      transitive: params.transitive ?? false,
    });
    return (payload as string[]) ?? [];
  }
}
