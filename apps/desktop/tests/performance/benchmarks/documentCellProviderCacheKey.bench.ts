import { LruCache, type CellData } from '@formula/grid/node';

import { DocumentCellProvider } from '../../../src/grid/shared/documentCellProvider.ts';

const CACHE_KEY_COL_STRIDE = 65_536;

type Bench = {
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
  clock?: 'wall' | 'cpu';
};

type CellState = { value: unknown; formula: string | null; styleId?: number };

/**
 * Baseline provider that mimics the previous shared-grid cache behavior:
 * `LruCache<string, CellData|null>` keyed by `${sheetId}:${row},${col}`.
 *
 * This is used for micro-benchmarking cache key generation overhead (string allocs)
 * vs the current numeric-key implementation.
 *
 * Enable via:
 *   FORMULA_BENCH_DOCUMENT_CELL_PROVIDER=1 pnpm benchmark
 */
class StringKeyDocumentCellProvider {
  cache = new LruCache<string, CellData | null>(50_000);
  options: {
    document: { getCell: (sheetId: string, coord: { row: number; col: number }) => CellState | null };
    getSheetId: () => string;
    headerRows: number;
    headerCols: number;
    rowCount: number;
    colCount: number;
  };

  constructor(options: {
    document: { getCell: (sheetId: string, coord: { row: number; col: number }) => CellState | null };
    getSheetId: () => string;
    headerRows: number;
    headerCols: number;
    rowCount: number;
    colCount: number;
  }) {
    this.options = options;
  }

  getCell(row: number, col: number): CellData | null {
    const { rowCount, colCount } = this.options;
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const sheetId = this.options.getSheetId();
    const key = `${sheetId}:${row},${col}`;
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;

    const docRow = row - this.options.headerRows;
    const docCol = col - this.options.headerCols;
    const state = this.options.document.getCell(sheetId, { row: docRow, col: docCol });
    if (!state) {
      this.cache.set(key, null);
      return null;
    }

    const cell: CellData = { row, col, value: state.value as any };
    this.cache.set(key, cell);
    return cell;
  }
}

function runGetCellLoop(provider: { getCell: (row: number, col: number) => unknown }, options: {
  startRow: number;
  startCol: number;
  viewportRows: number;
  viewportCols: number;
  frames: number;
}): number {
  let hash = 0;
  for (let frame = 0; frame < options.frames; frame++) {
    for (let r = 0; r < options.viewportRows; r++) {
      const row = options.startRow + r;
      for (let c = 0; c < options.viewportCols; c++) {
        const col = options.startCol + c;
        const cell = provider.getCell(row, col) as any;
        hash = ((hash << 5) - hash + (cell?.row ?? 0) + (cell?.col ?? 0)) | 0;
      }
    }
  }
  return hash;
}

export function createDocumentCellProviderCacheKeyBenchmarks(): Bench[] {
  const headerRows = 1;
  const headerCols = 1;
  const docRows = 1_048_576;
  const docCols = 16_384;

  const rowCount = docRows + headerRows;
  const colCount = docCols + headerCols;

  const viewportRows = 50;
  const viewportCols = 20;
  const frames = 10; // 50*20*10 = 10k getCell calls per benchmark iteration.

  // Avoid header cells (bench focuses on the cache-hit path, not header formatting).
  const startRow = headerRows + 10_000;
  const startCol = headerCols + 10;

  const doc = {
    getCell: (_sheetId: string, coord: { row: number; col: number }): CellState => ({
      value: coord.row * CACHE_KEY_COL_STRIDE + coord.col,
      formula: null,
      styleId: 0,
    }),
  };

  const numericProvider = new DocumentCellProvider({
    document: doc as any,
    getSheetId: () => 'Sheet1',
    headerRows,
    headerCols,
    rowCount,
    colCount,
    showFormulas: () => false,
    getComputedValue: () => null,
  });

  const stringProvider = new StringKeyDocumentCellProvider({
    document: doc,
    getSheetId: () => 'Sheet1',
    headerRows,
    headerCols,
    rowCount,
    colCount,
  });

  return [
    {
      name: 'grid.documentCellProvider.cacheKey.numeric.hit.p95',
      fn: () => {
        runGetCellLoop(numericProvider, { startRow, startCol, viewportRows, viewportCols, frames });
      },
      // Keep this lenient; this benchmark is primarily informational (dev-only).
      targetMs: 250,
      iterations: 20,
      warmup: 5,
    },
    {
      name: 'grid.documentCellProvider.cacheKey.string.hit.p95',
      fn: () => {
        runGetCellLoop(stringProvider, { startRow, startCol, viewportRows, viewportCols, frames });
      },
      // Keep this lenient; this benchmark is primarily informational (dev-only).
      targetMs: 750,
      iterations: 20,
      warmup: 5,
    },
  ];
}
