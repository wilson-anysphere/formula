export type RangeProvider = { getRange(rangeRef: string): any[][] };

export function createMatrixRangeProvider(sheets: Record<string, any[][]>): RangeProvider;

export function resolveSeries(
  chart: any,
  provider: RangeProvider
): Array<{ name: string | null; categories: any[]; values: any[]; xValues: any[]; yValues: any[] }>;

export function placeholderSvg(params: { width: number; height: number; label: string }): string;

export function renderChartSvg(chart: any, provider: RangeProvider, opts?: { width?: number; height?: number; theme?: any }): string;

export function renderChartSvgFromModel(
  model: any,
  liveData?: any,
  opts?: { width?: number; height?: number; theme?: any }
): Promise<string>;

