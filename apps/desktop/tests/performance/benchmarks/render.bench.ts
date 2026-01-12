type Grid = {
  rows: number;
  cols: number;
  data: string[];
};

function mulberry32(seed: number): () => number {
  return () => {
    // eslint-disable-next-line no-param-reassign
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function createGrid(rows: number, cols: number): Grid {
  const rand = mulberry32(1337);
  const data = new Array<string>(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      // Keep strings short but varied to emulate cell value formatting.
      const n = Math.floor(rand() * 1_000_000);
      data[r * cols + c] = n.toString(10);
    }
  }
  return { rows, cols, data };
}

function renderViewport(
  grid: Grid,
  startRow: number,
  startCol: number,
  viewportRows: number,
  viewportCols: number,
): number {
  // Simulates the critical path of a canvas grid frame:
  // - resolve visible cell data
  // - format strings + compute a cheap “measure text” proxy
  // - accumulate into a draw command list hash
  let hash = 0;
  for (let r = 0; r < viewportRows; r++) {
    const row = (startRow + r) % grid.rows;
    const base = row * grid.cols;
    for (let c = 0; c < viewportCols; c++) {
      const col = (startCol + c) % grid.cols;
      const text = grid.data[base + col]!;
      const width = text.length * 7;
      hash = ((hash << 5) - hash + width + row + col) | 0;
    }
  }
  return hash;
}

export function createRenderBenchmarks(): Array<{
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
  clock?: 'wall' | 'cpu';
}> {
  // Use a reasonably large backing grid to keep cache behavior realistic, while
  // benchmarking only the visible region.
  const grid = createGrid(10_000, 200);
  const viewportRows = 50;
  const viewportCols = 20;

  let scrollRow = 0;

  return [
    {
      name: 'render.frame.p95',
      fn: () => {
        renderViewport(grid, 0, 0, viewportRows, viewportCols);
      },
      // 60fps budget.
      targetMs: 16,
    },
    {
      name: 'render.scroll_step.p95',
      fn: () => {
        scrollRow = (scrollRow + 5) % grid.rows;
        renderViewport(grid, scrollRow, 0, viewportRows, viewportCols);
      },
      targetMs: 16,
    },
  ];
}
