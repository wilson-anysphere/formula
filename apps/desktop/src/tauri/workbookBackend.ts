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
  SheetUsedRange,
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
