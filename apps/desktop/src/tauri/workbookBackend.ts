export type SheetInfo = {
  id: string;
  name: string;
};

export type WorkbookInfo = {
  path: string | null;
  origin_path: string | null;
  sheets: SheetInfo[];
};

export type CellValue = {
  value: unknown | null;
  formula: string | null;
  display_value: string;
};

export type RangeData = {
  values: CellValue[][];
  start_row: number;
  start_col: number;
};

export type RangeCellEdit = {
  value: unknown | null;
  formula: string | null;
};

export type SheetUsedRange = {
  start_row: number;
  end_row: number;
  start_col: number;
  end_col: number;
};

export type WorkbookThemePalette = {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  followedHlink: string;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export class TauriWorkbookBackend {
  private readonly invoke: TauriInvoke;

  constructor() {
    this.invoke = getTauriInvoke();
  }

  async openWorkbook(path: string): Promise<WorkbookInfo> {
    const info = await this.invoke("open_workbook", { path });
    return info as WorkbookInfo;
  }

  async newWorkbook(): Promise<WorkbookInfo> {
    const info = await this.invoke("new_workbook");
    return info as WorkbookInfo;
  }

  async getWorkbookThemePalette(): Promise<WorkbookThemePalette | null> {
    const palette = await this.invoke("get_workbook_theme_palette");
    return (palette as WorkbookThemePalette | null) ?? null;
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
}
