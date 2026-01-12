import { JSDOM } from 'jsdom';

import { CanvasGridRenderer, type CellProvider } from '@formula/grid/node';

type BenchmarkDef = {
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
};

/**
 * CanvasGridRenderer is designed for browser/webview environments. Our perf suite
 * runs under Node, so we provide a minimal JSDOM + CanvasRenderingContext2D stub.
 *
 * The goal here is to benchmark the renderer's JS work (virtualization, text layout,
 * dirty region handling, scroll blitting decisions), not GPU/Skia time.
 */
function ensureDomAndCanvasMocks(): void {
  if (typeof window !== 'undefined' && typeof document !== 'undefined') return;

  const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    pretendToBeVisual: true,
    url: 'https://benchmark.local/',
  });

  const win = dom.window as unknown as Window & typeof globalThis;

  const defineGlobal = (key: string, value: unknown) => {
    const desc = Object.getOwnPropertyDescriptor(globalThis, key);
    if (desc && !desc.configurable) return;
    Object.defineProperty(globalThis, key, {
      value,
      configurable: true,
      writable: true,
    });
  };

  defineGlobal('window', win);
  defineGlobal('document', win.document);
  defineGlobal('HTMLElement', win.HTMLElement);
  defineGlobal('HTMLCanvasElement', win.HTMLCanvasElement);

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
  // Avoid per-call allocations: pick from a small pool of stable strings.
  const poolSize = 256;
  const values = new Array<string>(poolSize);
  for (let i = 0; i < poolSize; i++) {
    // Keep values short to avoid triggering text-overflow scan paths.
    values[i] = `v${i.toString(16).padStart(2, '0')}`;
  }

  return {
    getCell: (row, col) => {
      const idx = ((row * 31) ^ (col * 17)) & (poolSize - 1);
      return { row, col, value: values[idx]! };
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

  // Reuse a single renderer instance for scroll benchmarking to avoid mixing in setup cost.
  const scrollRenderer = new CanvasGridRenderer({ provider, rowCount, colCount });
  scrollRenderer.setPerfStatsEnabled(false);
  scrollRenderer.attach({
    grid: document.createElement('canvas'),
    content: document.createElement('canvas'),
    selection: document.createElement('canvas'),
  });
  scrollRenderer.setFrozen(frozenRows, frozenCols);
  scrollRenderer.resize(viewportWidth, viewportHeight, devicePixelRatio);
  scrollRenderer.renderImmediately();

  const deltaY = 21 * 5; // 5 rows per "wheel" step at default 21px row height.

  return [
    {
      name: 'gridRenderer.firstFrame.p95',
      fn: () => {
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
        renderer.destroy();
      },
      // 60fps budget.
      targetMs: 16,
      iterations: 25,
      warmup: 5,
    },
    {
      name: 'gridRenderer.scrollStep.p95',
      fn: () => {
        scrollRenderer.scrollBy(0, deltaY);
        scrollRenderer.renderImmediately();
      },
      targetMs: 16,
    },
  ];
}
