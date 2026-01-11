export type WrapMode = "none" | "word" | "char";
export type TextDirection = "ltr" | "rtl" | "auto";
export type TextAlign = "left" | "right" | "center" | "start" | "end";

export type FontSpec = {
  family: string;
  sizePx: number;
  weight?: string | number;
  style?: string;
};

export type TextMeasurement = {
  width: number;
  ascent: number;
  descent: number;
};

export interface TextMeasurer {
  measure(text: string, font: FontSpec): TextMeasurement;
}

export type TextRun = {
  text: string;
  font?: FontSpec;
  color?: string;
  underline?: boolean;
};

export type LayoutOptions = {
  text?: string;
  runs?: TextRun[];
  font: FontSpec;
  maxWidth: number;
  wrapMode: WrapMode;
  align: TextAlign;
  direction?: TextDirection;
  lineHeightPx?: number;
  maxLines?: number;
  ellipsis?: string;
  locale?: string;
};

export type LayoutLine = {
  text: string;
  runs: TextRun[];
  width: number;
  ascent: number;
  descent: number;
  /** X offset from the layout origin. */
  x: number;
};

export type TextLayout = {
  lines: LayoutLine[];
  width: number;
  height: number;
  lineHeight: number;
  direction: "ltr" | "rtl";
  maxWidth: number;
  resolvedAlign: "left" | "right" | "center";
};

export class TextLayoutEngine {
  constructor(
    measurer: TextMeasurer,
    opts?: { maxMeasureCacheEntries?: number; maxLayoutCacheEntries?: number },
  );
  measure(text: string, font: FontSpec): TextMeasurement;
  layout(options: LayoutOptions): TextLayout;
}

export class CanvasTextMeasurer implements TextMeasurer {
  constructor(ctx: CanvasRenderingContext2D);
  measure(text: string, font: FontSpec): TextMeasurement;
}

export function createCanvasTextMeasurer(): CanvasTextMeasurer;

export type HarfBuzzInstance = any;

export type HarfBuzzFontFace = {
  key: string;
  family: string;
  weight: string | number;
  style: string;
  upem: number;
  ascentRatio: number;
  descentRatio: number;
};

export class HarfBuzzFontManager {
  constructor(hb: HarfBuzzInstance);
  /** Incremented when font data or fallback configuration changes. */
  version: number;
  /** Underlying HarfBuzz instance. */
  hb: HarfBuzzInstance;
  loadFont(data: ArrayBuffer | ArrayBufferView, spec: Omit<FontSpec, "sizePx">): HarfBuzzFontFace;
  setFallbackFamilies(families: string[]): void;
  getFace(spec: FontSpec | Omit<FontSpec, "sizePx">): HarfBuzzFontFace;
  getFallbackFaces(spec: FontSpec | Omit<FontSpec, "sizePx">): HarfBuzzFontFace[];
  destroy(): void;
}

export class HarfBuzzTextMeasurer implements TextMeasurer {
  constructor(fontManager: HarfBuzzFontManager, opts?: { maxShapeCacheEntries?: number });
  fontManager: HarfBuzzFontManager;
  measure(text: string, font: FontSpec): TextMeasurement;
}

export function loadHarfBuzz(): Promise<HarfBuzzInstance>;

export function createHarfBuzzTextMeasurer(opts?: {
  fonts?: Array<Omit<FontSpec, "sizePx"> & { data: ArrayBuffer | ArrayBufferView }>;
  fallbackFamilies?: string[];
  maxShapeCacheEntries?: number;
}): Promise<HarfBuzzTextMeasurer>;

export function drawTextLayout(
  ctx: CanvasRenderingContext2D,
  layout: TextLayout,
  x: number,
  y: number,
  opts?: { rotationRad?: number },
): void;

export function detectBaseDirection(text: string): "ltr" | "rtl";

export function resolveAlign(
  align: TextAlign,
  direction: "ltr" | "rtl",
): "left" | "right" | "center";

export function normalizeFont(font: FontSpec): Required<
  Pick<FontSpec, "family" | "sizePx">
> & { weight: string | number; style: string };

export function fontKey(font: FontSpec): string;

export function toCanvasFontString(font: FontSpec): string;
