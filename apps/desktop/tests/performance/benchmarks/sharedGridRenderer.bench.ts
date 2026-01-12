import { JSDOM } from 'jsdom';

import { CanvasGridRenderer, type CellData, type CellProvider } from '@formula/grid/node';

type BenchmarkDef = {
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
  clock?: 'wall' | 'cpu';
};

/**
 * CanvasGridRenderer is designed for browser/webview environments. Our perf suite
 * runs under Node, so we provide a minimal JSDOM + CanvasRenderingContext2D stub.
 *
 * The goal here is to benchmark the renderer's JS work (virtualization, text layout,
 * dirty region handling, scroll blitting decisions), not GPU/Skia time.
 */
function ensureDomAndCanvasMocks(): void {
  const defineGlobal = (key: string, value: unknown) => {
    const desc = Object.getOwnPropertyDescriptor(globalThis, key);
    if (desc && !desc.configurable) return;
    Object.defineProperty(globalThis, key, {
      value,
      configurable: true,
      writable: true,
    });
  };

  let win: Window & typeof globalThis;
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    const dom = new JSDOM('<!doctype html><html><body></body></html>', {
      pretendToBeVisual: true,
      url: 'https://benchmark.local/',
    });
    win = dom.window as unknown as Window & typeof globalThis;

    defineGlobal('window', win);
    defineGlobal('document', win.document);
    defineGlobal('HTMLElement', win.HTMLElement);
    defineGlobal('HTMLCanvasElement', win.HTMLCanvasElement);
  } else {
    win = window as unknown as Window & typeof globalThis;
  }

  // The renderer schedules frames via rAF when state changes. Benchmarks drive frames
  // manually via `renderImmediately()`, so keep rAF a cheap no-op.
  defineGlobal('requestAnimationFrame', () => 0);

  const noop = () => {};
  const sharedMetrics = {
    width: 0,
    actualBoundingBoxAscent: 8,
    actualBoundingBoxDescent: 2,
  } as TextMetrics;

  const createMock2dContext = (canvas: HTMLCanvasElement): CanvasRenderingContext2D =>
    ({
      canvas,
      fillStyle: '#000',
      strokeStyle: '#000',
      lineWidth: 1,
      font: '',
      textAlign: 'left',
      textBaseline: 'alphabetic',
      globalAlpha: 1,
      imageSmoothingEnabled: false,
      setTransform: noop,
      clearRect: noop,
      fillRect: noop,
      strokeRect: noop,
      beginPath: noop,
      rect: noop,
      clip: noop,
      fill: noop,
      stroke: noop,
      moveTo: noop,
      lineTo: noop,
      closePath: noop,
      save: noop,
      restore: noop,
      drawImage: noop,
      translate: noop,
      rotate: noop,
      setLineDash: noop,
      fillText: noop,
      measureText: (text: string) => {
        // Keep deterministic and cheap; renderer/text-layout only needs width + basic metrics.
        //
        // We reuse a single TextMetrics-like object to avoid allocations/GC noise in the
        // benchmark itself (the renderer consumes the fields synchronously).
        (sharedMetrics as unknown as { width: number }).width = text.length * 6;
        return sharedMetrics;
      },
    }) as unknown as CanvasRenderingContext2D;

  const ctxCache = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
  win.HTMLCanvasElement.prototype.getContext = function getContextStub() {
    const canvas = this as HTMLCanvasElement;
    const cached = ctxCache.get(canvas);
    if (cached) return cached;
    const created = createMock2dContext(canvas);
    ctxCache.set(canvas, created);
    return created;
  } as unknown as typeof win.HTMLCanvasElement.prototype.getContext;
}

function createDeterministicProvider(): CellProvider {
  // Avoid per-call allocations: use a small pool of stable strings and reuse CellData
  // objects so GC pauses don't dominate benchmark p95.
  const poolSize = 256;
  const values = new Array<string>(poolSize);
  for (let i = 0; i < poolSize; i++) {
    // Keep values short to avoid triggering text-overflow scan paths.
    values[i] = `v${i.toString(16).padStart(2, '0')}`;
  }

  const cells: CellData[] = values.map((value) => ({
    // The renderer does not depend on the returned cell's `row`/`col` fields (it
    // uses the indices passed to `getCell`), so keep them constant to allow
    // object reuse across the entire grid.
    row: 0,
    col: 0,
    value,
  }));

  return {
    getCell: (row, col) => {
      const idx = ((row * 31) ^ (col * 17)) & (poolSize - 1);
      return cells[idx] ?? null;
    },
  };
}

export function createSharedGridRendererBenchmarks(): BenchmarkDef[] {
  ensureDomAndCanvasMocks();

  // Excel-ish maxes keep integer math realistic while still rendering only the viewport.
  const rowCount = 1_048_576;
  const colCount = 16_384;
  const viewportWidth = 1200;
  const viewportHeight = 700;
  const frozenRows = 1;
  const frozenCols = 1;
  const devicePixelRatio = 1;

  const provider = createDeterministicProvider();

  const createInitializedRenderer = (): CanvasGridRenderer => {
    const renderer = new CanvasGridRenderer({ provider, rowCount, colCount });
    renderer.setPerfStatsEnabled(false);
    renderer.attach({
      grid: document.createElement('canvas'),
      content: document.createElement('canvas'),
      selection: document.createElement('canvas'),
    });
    renderer.setFrozen(frozenRows, frozenCols);
    renderer.resize(viewportWidth, viewportHeight, devicePixelRatio);
    renderer.renderImmediately();
    return renderer;
  };

  // Lazily create renderers so `gridRenderer.firstFrame.p95` measures in isolation.
  // The scroll benchmarks have warmup iterations, so the one-time init cost is absorbed there.
  let scrollRendererY: CanvasGridRenderer | null = null;
  let scrollRendererX: CanvasGridRenderer | null = null;

  const deltaY = 21 * 5; // 5 rows per "wheel" step at default 21px row height.
  const deltaX = 100 * 3; // 3 columns per step at default 100px col width.

  // Reuse DOM canvases across first-frame iterations to avoid benchmarking JSDOM element allocation
  // rather than renderer work.
  const firstFrameCanvases = {
    grid: document.createElement('canvas'),
    content: document.createElement('canvas'),
    selection: document.createElement('canvas'),
  };

  const benchmarks: BenchmarkDef[] = [
    {
      name: 'gridRenderer.firstFrame.p95',
      fn: () => {
        const renderer = new CanvasGridRenderer({ provider, rowCount, colCount });
        renderer.setPerfStatsEnabled(false);
        renderer.attach(firstFrameCanvases);
        renderer.setFrozen(frozenRows, frozenCols);
        renderer.resize(viewportWidth, viewportHeight, devicePixelRatio);
        renderer.renderImmediately();
        renderer.destroy();
      },
      // 60fps budget.
      targetMs: 16,
      clock: 'cpu',
      // Use enough samples that occasional Node/V8 GC pauses don't dominate p95.
      iterations: 100,
      warmup: 10,
    },
    {
      name: 'gridRenderer.scrollStep.p95',
      fn: () => {
        const renderer = (scrollRendererY ??= createInitializedRenderer());
        renderer.scrollBy(0, deltaY);
        renderer.renderImmediately();
      },
      targetMs: 16,
      clock: 'cpu',
      // Use enough samples that occasional Node/V8 GC pauses don't dominate p95.
      iterations: 200,
      warmup: 20,
    },
    {
      name: 'gridRenderer.scrollStepHorizontal.p95',
      fn: () => {
        const renderer = (scrollRendererX ??= createInitializedRenderer());
        renderer.scrollBy(deltaX, 0);
        renderer.renderImmediately();
      },
      targetMs: 16,
      clock: 'cpu',
      // Use enough samples that occasional Node/V8 GC pauses don't dominate p95.
      iterations: 200,
      warmup: 20,
    },
  ];

  // Batched axis override application can be quite noisy under Node (sorting + Map churn + GC),
  // which risks making the overall perf suite flaky. Keep it opt-in.
  if (process.env.FORMULA_BENCH_GRID_AXIS_OVERRIDES === '1') {
    // Separate renderer for axis override benchmarking so we don't perturb the scroll benchmark.
    const axisRenderer = new CanvasGridRenderer({ provider, rowCount, colCount });
    axisRenderer.setPerfStatsEnabled(false);
    axisRenderer.attach({
      grid: document.createElement('canvas'),
      content: document.createElement('canvas'),
      selection: document.createElement('canvas'),
    });
    axisRenderer.setFrozen(frozenRows, frozenCols);
    axisRenderer.resize(viewportWidth, viewportHeight, devicePixelRatio);
    axisRenderer.renderImmediately();

    const overrideRows = 10_000;
    const overrideCols = 5_000;

    const axisRowOverridesA = new Map<number, number>();
    const axisRowOverridesB = new Map<number, number>();
    for (let i = 0; i < overrideRows; i++) {
      const row = i * 3;
      axisRowOverridesA.set(row, 30);
      axisRowOverridesB.set(row, 31);
    }

    const axisColOverridesA = new Map<number, number>();
    const axisColOverridesB = new Map<number, number>();
    for (let i = 0; i < overrideCols; i++) {
      const col = i * 2;
      axisColOverridesA.set(col, 120);
      axisColOverridesB.set(col, 121);
    }

    let axisOverrideToggle = false;

    benchmarks.push({
      name: 'gridRenderer.applyAxisSizeOverrides.10k.p95',
      fn: () => {
        // Toggle between two override maps so every iteration performs a real update (vs the
        // renderer short-circuiting when the override set is unchanged).
        axisOverrideToggle = !axisOverrideToggle;

        axisRenderer.applyAxisSizeOverrides(
          {
            rows: axisOverrideToggle ? axisRowOverridesA : axisRowOverridesB,
            cols: axisOverrideToggle ? axisColOverridesA : axisColOverridesB,
          },
          { resetUnspecified: true },
        );
      },
      // Keep targets conservative: GitHub Actions runners vary, but this should still catch
      // pathological regressions (e.g. reintroducing O(n^2) axis updates).
      targetMs: 120,
      iterations: 25,
      warmup: 5,
    });
  }

  return benchmarks;
}
