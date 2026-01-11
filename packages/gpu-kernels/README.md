# @formula/gpu-kernels
Optional WebGPU compute acceleration for heavy spreadsheet/analytics kernels with CPU fallback.

## What’s implemented
- **SUM / SUMPRODUCT** (parallel reduction)
- **MIN / MAX / AVERAGE / COUNT** (reductions; AVERAGE derived from SUM)
- **Group-by aggregations** on key columns:
  - `COUNT`, `SUM`, `MIN`, `MAX`
  - Keys: `Uint32Array` (dictionary ids) or `Int32Array` (signed 32-bit keys)
  - Outputs are returned sorted by `uniqueKeys`
  - Two-key CPU group-by variants: `groupByCount2/groupBySum2/groupByMin2/groupByMax2`
- **Hash join** on two key arrays producing matching `(leftIndex, rightIndex)` pairs (inner and left join)
- **MMULT** (matrix multiplication)
- **Sort** (bitonic sort for numeric vectors)
- **Histogram/binning** (atomic bin counts)

All kernels have a **CPU fallback** and will run without WebGPU.

## Usage (high-level engine)
```js
import { createKernelEngine } from "./src/index.js";

const engine = await createKernelEngine({
  // Default is "excel" (compat-safe): never silently downcasts f64->f32.
  precision: "excel",
  gpu: {
    enabled: true
  }
});

const values = new Float64Array(1_000_000).fill(1);
const sum = await engine.sum(values);
console.log(sum);

// Group-by SUM(+COUNT).
const keys = new Uint32Array([1, 1, 2, 2, 2]);
const valuesByKey = new Float64Array([10, 5, 1, 2, 3]);
const grouped = await engine.groupBySum(keys, valuesByKey);
console.log(grouped.uniqueKeys, grouped.sums, grouped.counts);

// Two-key group-by SUM(+COUNT) (CPU-only for now).
const keysA = new Uint32Array([1, 1, 2, 2]);
const keysB = new Uint32Array([10, 11, 10, 11]);
const grouped2 = await engine.groupBySum2(keysA, keysB, new Float64Array([1, 2, 3, 4]));
console.log(grouped2.uniqueKeysA, grouped2.uniqueKeysB, grouped2.sums, grouped2.counts);

// Hash join (inner join).
const leftKeys = new Uint32Array([1, 2, 2, 3]);
const rightKeys = new Uint32Array([2, 2, 3, 4]);
const joined = await engine.hashJoin(leftKeys, rightKeys);
console.log(joined.leftIndex, joined.rightIndex);

console.log(engine.diagnostics());
await engine.dispose();
```

## WebGPU notes
- WebGPU is optional. If `navigator.gpu` is unavailable or device creation fails, the engine falls back to CPU.
- When the adapter supports the `shader-f64` feature, the backend compiles **both f32 and f64 WGSL pipelines**.
- Precision modes:
  - `precision: "excel"` (default): uses **f64** GPU kernels when supported and otherwise falls back to CPU. This mode never silently downcasts `Float64Array` inputs to `f32`.
  - `precision: "fast"`: prefers **f32** GPU kernels and may downcast `Float64Array` inputs for performance.
- Current WebGPU coverage:
  - `groupBySum/groupByMin/groupByMax` are currently **f32-only** (implemented via atomic CAS loops), so `"excel"` mode will fall back to CPU for these.
  - `groupBy*2` (two-key group-by) is currently **CPU-only**.
- Optional safety net: in `"excel"` mode the engine can validate some GPU results against CPU for smaller workloads (default `maxElements=32768`) and fall back to CPU if the difference exceeds a strict tolerance (`abs=1e-9`, `rel=1e-12`). Configure via the `validation` option.
- In `"excel"` mode, if a WebGPU kernel throws at runtime (pipeline/dispatch/device issues), the engine records diagnostics and **falls back to CPU**.
- Kernel edge cases:
  - **Sort** matches `TypedArray#sort` semantics for special values: `NaN` sorts to the end, and `±Infinity` sorts normally.
  - **Histogram** ignores `NaN` values and clamps `±Infinity` into the first/last bin.
  - **Group-by SUM/MIN/MAX** follow JS numeric semantics (`NaN` propagates, `±Infinity` behaves per IEEE-754). `MIN/MAX` preserve signed zero like `Math.min/Math.max`.
  - **Hash join** supports:
    - `joinType: "inner"` (default): returns *all* matching pairs (duplicates produce multiple output rows)
    - `joinType: "left"`: also includes unmatched left rows with `rightIndex=0xFFFF_FFFF`
    - Outputs are sorted by `(leftIndex, rightIndex)` ascending
  - **Hash join** requires both key arrays to use the same numeric domain (`Int32Array` with `Int32Array`, or `Uint32Array` with `Uint32Array`).

## Benchmarks
```bash
pnpm bench:gpu-kernels
```

## Diagnostics UI (example)
Serve the repo with any static server (required for module imports), then open:
`packages/gpu-kernels/examples/diagnostics.html`
