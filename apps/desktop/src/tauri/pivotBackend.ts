import type { SheetInfo, TableInfo } from "./workbookBackend";

export type PivotCellRange = {
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
};

export type PivotDestination = {
  sheet_id: string;
  row: number;
  col: number;
};

export type CellUpdate = {
  sheet_id: string;
  row: number;
  col: number;
  value: unknown | null;
  formula: string | null;
  display_value: string;
};

export type CreatePivotTableRequest = {
  name: string;
  source_sheet_id: string;
  source_range: PivotCellRange;
  destination: PivotDestination;
  // Rust serde expects camelCase keys inside the config object.
  config: Record<string, unknown>;
};

export type CreatePivotTableResponse = {
  pivot_id: string;
  updates: CellUpdate[];
};

export type PivotTableSummary = {
  id: string;
  name: string;
  source_sheet_id: string;
  source_range: PivotCellRange;
  destination: PivotDestination;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export class TauriPivotBackend {
  private readonly invoke: TauriInvoke;

  constructor(options: { invoke?: TauriInvoke } = {}) {
    this.invoke = options.invoke ?? getTauriInvoke();
  }

  async addSheet(name: string): Promise<SheetInfo> {
    const payload = await this.invoke("add_sheet", { name });
    return payload as SheetInfo;
  }

  async listTables(): Promise<TableInfo[]> {
    const payload = await this.invoke("list_tables");
    return (payload as TableInfo[]) ?? [];
  }

  async listPivotTables(): Promise<PivotTableSummary[]> {
    const payload = await this.invoke("list_pivot_tables");
    return (payload as PivotTableSummary[]) ?? [];
  }

  async createPivotTable(request: CreatePivotTableRequest): Promise<CreatePivotTableResponse> {
    const payload = await this.invoke("create_pivot_table", { request });
    return payload as CreatePivotTableResponse;
  }

  async refreshPivotTable(pivotId: string): Promise<CellUpdate[]> {
    const payload = await this.invoke("refresh_pivot_table", { request: { pivot_id: pivotId } });
    return (payload as CellUpdate[]) ?? [];
  }
}

