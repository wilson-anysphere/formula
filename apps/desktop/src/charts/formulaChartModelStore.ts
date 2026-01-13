import type { ChartStore } from "./chartRendererAdapter";
import { formulaModelChartModelToUiChartModel } from "./formulaModelChartModel.js";
import type { ChartModel, ChartTheme, ResolvedChartData } from "./renderChart";

/**
 * In-memory chart model store for charts imported from a workbook (e.g. XLSX DrawingML).
 *
 * The keys are expected to be globally unique across the workbook. A recommended
 * pattern is `${sheetId}:${drawingObjectId}` (stable across sheet renames).
 */
export class FormulaChartModelStore implements ChartStore {
  private readonly models = new Map<string, ChartModel>();
  private readonly data = new Map<string, Partial<ResolvedChartData>>();
  private readonly themes = new Map<string, Partial<ChartTheme>>();
  private defaultTheme: Partial<ChartTheme> | undefined;

  static chartIdFromSheetObject(sheetId: string, drawingObjectId: string | number): string {
    return `${sheetId}:${String(drawingObjectId)}`;
  }

  static chartIdFromDrawingRel(drawingPart: string, relId: string): string {
    return `${drawingPart}:${relId}`;
  }

  getChartModel(chartId: string): ChartModel | undefined {
    return this.models.get(chartId);
  }

  getChartData(chartId: string): Partial<ResolvedChartData> | undefined {
    return this.data.get(chartId);
  }

  getChartTheme(chartId: string): Partial<ChartTheme> | undefined {
    return this.themes.get(chartId) ?? this.defaultTheme;
  }

  setChartModel(chartId: string, model: ChartModel): void {
    this.models.set(chartId, model);
  }

  /**
   * Convenience helper for backends that expose the Rust `ChartModel` via JSON.
   */
  setFormulaModelChartModel(chartId: string, formulaModel: unknown): void {
    const converted = formulaModelChartModelToUiChartModel(formulaModel as any) as unknown as ChartModel;
    this.models.set(chartId, converted);
  }

  setChartData(chartId: string, data: Partial<ResolvedChartData> | undefined): void {
    if (!data) this.data.delete(chartId);
    else this.data.set(chartId, data);
  }

  setChartTheme(chartId: string, theme: Partial<ChartTheme> | undefined): void {
    if (!theme) this.themes.delete(chartId);
    else this.themes.set(chartId, theme);
  }

  /**
   * Set a theme patch applied to all charts when no per-chart theme is present.
   *
   * This is convenient for aligning imported chart rendering with the workbook's
   * theme palette without having to enumerate every chart id.
   */
  setDefaultTheme(theme: Partial<ChartTheme> | undefined): void {
    this.defaultTheme = theme;
  }

  clear(): void {
    this.models.clear();
    this.data.clear();
    this.themes.clear();
  }
}
