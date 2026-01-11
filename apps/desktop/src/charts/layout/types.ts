export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ViewportRect extends Rect {}

export interface LineSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface FontSpec {
  family: string;
  sizePx: number;
  weight?: number | string;
  style?: "normal" | "italic";
}

export interface ChartTheme {
  fonts: {
    title: FontSpec;
    axis: FontSpec;
    legend: FontSpec;
  };
}

export type ChartKind = "bar" | "line" | "pie" | "scatter" | "unknown";

export interface ChartTypeModel {
  kind: ChartKind;
  name?: string | null;
}

export type ChartDataCache<T> =
  | { cache: Array<T | null> }
  | { cache?: Array<T | null>; ref?: string | null }
  | Array<T | null>;

export interface ChartSeriesModel {
  name?: string | null;
  categories?: ChartDataCache<string | number> | null;
  values?: ChartDataCache<number | string> | null;
  xValues?: ChartDataCache<number | string> | null;
  yValues?: ChartDataCache<number | string> | null;
}

export interface ChartLegendModel {
  /**
   * v1 layout only supports legend on the right (or none), but we accept
   * OOXML-style single-letter positions as input ("r") since the Rust parser
   * may preserve them.
   */
  position?: "right" | "none" | "r" | "l" | "t" | "b" | null;
  overlay?: boolean | null;
}

export interface AxisScalingModel {
  min?: number | null;
  max?: number | null;
  reverseOrder?: boolean | null;
  /**
   * Optional OOXML scaling orientation ("minMax" or "maxMin"). If present, it
   * should be treated as the source of truth for reverse order behavior.
   */
  orientation?: "minMax" | "maxMin" | null;
}

export interface ChartAxisModel {
  /**
   * User-visible axis kind (category vs numeric value axis).
   */
  kind: "category" | "value" | "catAx" | "valAx";
  /**
   * Axis position. Accept both human-friendly values and OOXML-style single
   * letters (b/l/r/t).
   */
  position: "left" | "right" | "top" | "bottom" | "l" | "r" | "t" | "b";
  id?: string | null;
  scaling?: AxisScalingModel | null;
  /**
   * Presence (truthy) means the axis should render major gridlines.
   */
  majorGridlines?: unknown;
  /**
   * Excel format code for tick labels (c:numFmt/@formatCode).
   */
  formatCode?: string | null;
  /**
   * Alias for formatCode used by some model variants.
   */
  numberFormatCode?: string | null;
}

export interface ChartModel {
  chartType: ChartTypeModel;
  title?: string | null;
  legend?: ChartLegendModel | null;
  axes?: ChartAxisModel[] | null;
  series: ChartSeriesModel[];
}

export interface TextLineLayout {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TextBlockLayout {
  rect: Rect;
  font: FontSpec;
  lineHeight: number;
  align: "start" | "center" | "end";
  lines: TextLineLayout[];
}

export interface LegendEntryLayout {
  seriesIndex: number;
  label: string;
  markerRect: Rect;
  labelRect: Rect;
}

export interface LegendLayout {
  rect: Rect;
  font: FontSpec;
  entries: LegendEntryLayout[];
}

export interface TickLayout {
  value: number | string;
  label: string;
  position: { x: number; y: number };
  labelRect: Rect;
}

export interface AxisLayout {
  id: string;
  orientation: "x" | "y";
  kind: "category" | "value";
  axisLine: LineSegment;
  ticks: TickLayout[];
  gridlines: LineSegment[];
}

export interface LinearScale {
  type: "linear";
  domain: [number, number];
  range: [number, number];
}

export interface BandScale {
  type: "band";
  domain: string[];
  range: [number, number];
  step: number;
  bandwidth: number;
}

export type Scale = LinearScale | BandScale;

export interface ChartLayout {
  chartAreaRect: Rect;
  plotAreaRect: Rect;

  titleRect: Rect | null;
  titleText: TextBlockLayout | null;

  legendRect: Rect | null;
  legend: LegendLayout | null;

  axes: Record<string, AxisLayout>;
  scales: Record<string, Scale>;
}
