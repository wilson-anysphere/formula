import type { EngineClient } from "@formula/engine";
import { EngineCellCache, EngineGridProvider } from "@formula/spreadsheet-frontend";

export interface EngineCellProviderOptions {
  engine: EngineClient;
  rowCount: number;
  colCount: number;
  sheet?: string;
  cache?: EngineCellCache;
}

/**
 * Thin web-preview adapter around the shared `EngineGridProvider` that owns a
 * cache instance by default so `App` can stay focused on initialization/demo
 * state.
 */
export class EngineCellProvider extends EngineGridProvider {
  readonly cache: EngineCellCache;

  constructor(options: EngineCellProviderOptions) {
    const cache = options.cache ?? new EngineCellCache(options.engine);
    super({ cache, rowCount: options.rowCount, colCount: options.colCount, sheet: options.sheet, headers: true });
    this.cache = cache;
  }
}

