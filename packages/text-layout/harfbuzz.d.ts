import type { FontSpec, TextMeasurement, TextMeasurer } from ".";

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
  cacheKey: string;
  measure(text: string, font: FontSpec): TextMeasurement;
}

export function loadHarfBuzz(): Promise<HarfBuzzInstance>;

export function createHarfBuzzTextMeasurer(opts?: {
  fonts?: Array<Omit<FontSpec, "sizePx"> & { data: ArrayBuffer | ArrayBufferView }>;
  fallbackFamilies?: string[];
  maxShapeCacheEntries?: number;
}): Promise<HarfBuzzTextMeasurer>;
