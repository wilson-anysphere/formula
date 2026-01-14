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

import type { ImportedEmbeddedCellImage as ImportedEmbeddedCellImageInfo } from "../workbook/load/embeddedCellImages.js";

import { getTauriInvokeOrThrow, type TauriInvoke } from "./api";
export type SheetViewOverrides = {
  /**
   * Sparse column width overrides in Excel "character" units, keyed by 0-based column index.
   *
   * Note: These widths still need to be converted into the UI's view units (px) before being
   * stored in DocumentController sheet view state.
   */
  colWidths?: Record<string, number>;
  /**
   * 0-based indices for user-hidden columns.
   */
  hiddenCols?: number[];
};

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

export type { ImportedEmbeddedCellImageInfo };

export type ImportedSheetBackgroundImageInfo = {
  sheet_name: string;
  worksheet_part: string;
  image_id: string;
  bytes_base64: string;
  mime_type: string;
};

export class TauriWorkbookBackend implements WorkbookBackend {
  private readonly invoke: TauriInvoke;

  constructor() {
    this.invoke = getTauriInvokeOrThrow();
  }

  async newWorkbook(): Promise<WorkbookInfo> {
    const info = await this.invoke("new_workbook");
    return info as WorkbookInfo;
  }

  async openWorkbook(path: string, options?: { password?: string }): Promise<WorkbookInfo> {
    const args: Record<string, unknown> = { path };
    if (options?.password != null) args.password = options.password;
    const info = await this.invoke("open_workbook", args);
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

  /**
   * Desktop-only: fetch embedded-in-cell images parsed from the opened XLSX package.
   *
   * These correspond to Excel RichData (`vm=`) cell images ("Place in Cell" / `IMAGE()`).
   */
  async listImportedEmbeddedCellImages(): Promise<ImportedEmbeddedCellImageInfo[]> {
    const payload = await this.invoke("list_imported_embedded_cell_images");
    return (payload as ImportedEmbeddedCellImageInfo[]) ?? [];
  }

  /**
   * Desktop-only: fetch DrawingML objects + image bytes parsed from the opened XLSX package.
   *
   * Payload shape (best-effort; callers should tolerate missing fields):
   * - `{ drawings: [{ sheet_name, sheet_part, drawing_part, objects: DrawingObject[] }], images: [{ id, bytesBase64, mimeType }] }`
   */
  async getImportedDrawingLayer(): Promise<{ drawings: unknown[]; images: unknown[] }> {
    const payload = await this.invoke("list_imported_drawing_objects");
    const p = payload as any;
    const drawings = Array.isArray(p?.drawings) ? p.drawings : [];
    const images = Array.isArray(p?.images) ? p.images : [];
    return { drawings, images };
  }

  /**
   * Desktop-only: fetch worksheet background images (`<picture r:id="...">`) extracted from the
   * opened XLSX package.
   */
  async listImportedSheetBackgroundImages(): Promise<ImportedSheetBackgroundImageInfo[]> {
    const payload = await this.invoke("list_imported_sheet_background_images");
    return (payload as ImportedSheetBackgroundImageInfo[]) ?? [];
  }

  async saveWorkbook(path?: string, options?: { password?: string }): Promise<void> {
    const args: Record<string, unknown> = {};
    if (path) args.path = path;
    if (options?.password != null) args.password = options.password;
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
   * Best-effort: older desktop backends may not expose sheet view overrides.
   */
  async getSheetViewOverrides(sheetId: string): Promise<SheetViewOverrides | null> {
    const payload = await this.invoke("get_sheet_view_overrides", {
      sheet_id: sheetId,
    });
    return (payload as SheetViewOverrides | null) ?? null;
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

  /**
   * Best-effort: not all backend builds expose sheet view snapshot commands.
   * Callers should tolerate failures and treat them as "no persisted view state".
   */
  async getSheetViewState(sheetId: string): Promise<unknown | null> {
    const payload = await this.invoke("get_sheet_view_state", {
      sheet_id: sheetId,
    });
    return (payload as unknown | null) ?? null;
  }

  /**
   * Desktop-only: fetch worksheet column properties imported from the source workbook model.
   *
   * This is distinct from persisted UI sheet view state (`formula_ui_sheet_view`): it reflects
   * the workbook's native OOXML `<cols>` metadata (`col/@width` + `col/@hidden`).
   *
   * Best-effort: not all backend builds expose this command. Callers should tolerate failures
   * and treat them as "no imported column metadata".
   */
  async getSheetImportedColProperties(sheetId: string): Promise<unknown | null> {
    const payload = await this.invoke("get_sheet_imported_col_properties", {
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
