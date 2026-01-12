import { createRenderBenchmarks } from './render.bench.ts';

export function createStartupBenchmarks(): Array<{
  name: string;
  fn: () => void;
  targetMs: number;
  iterations?: number;
  warmup?: number;
  clock?: 'wall' | 'cpu';
}> {
  const [renderFrame] = createRenderBenchmarks();

  return [
    {
      name: 'startup.bootstrap_to_first_render.p95',
      fn: () => {
        // Simulated “cold start to interactive” for the JS layer:
        // - allocate model structures
        // - build initial view caches
        // - render first frame
        //
        // This intentionally keeps work modest for now and acts as a regression
        // guardrail as real startup code is introduced.
        const rows = 2_000;
        const cols = 50;
        const data = new Array<number>(rows * cols);
        for (let i = 0; i < data.length; i++) data[i] = i ^ (i >>> 3);

        // Basic indexing cache.
        const rowOffsets = new Array<number>(rows);
        for (let r = 0; r < rows; r++) rowOffsets[r] = r * cols;

        // Touch some values to avoid dead-code elimination.
        let checksum = 0;
        for (let r = 0; r < 200; r++) checksum ^= data[rowOffsets[r]!];
        if (checksum === 42) throw new Error('unreachable');

        renderFrame!.fn();
      },
      targetMs: 1000,
      iterations: 30,
      warmup: 5,
    },
  ];
}
