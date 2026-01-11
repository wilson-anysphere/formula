# @formula/gpu-kernels
Optional WebGPU compute acceleration for heavy spreadsheet/analytics kernels with CPU fallback.

## What’s implemented
- **SUM / SUMPRODUCT** (parallel reduction)
- **MIN / MAX / AVERAGE / COUNT** (reductions; AVERAGE derived from SUM)
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
console.log(engine.diagnostics());
await engine.dispose();
```

## WebGPU notes
- WebGPU is optional. If `navigator.gpu` is unavailable or device creation fails, the engine falls back to CPU.
- When the adapter supports the `shader-f64` feature, the backend compiles **both f32 and f64 WGSL pipelines**.
- Precision modes:
  - `precision: "excel"` (default): uses **f64** GPU kernels when supported and otherwise falls back to CPU. This mode never silently downcasts `Float64Array` inputs to `f32`.
  - `precision: "fast"`: prefers **f32** GPU kernels and may downcast `Float64Array` inputs for performance.
- Optional safety net: in `"excel"` mode the engine can validate some GPU results against CPU for smaller workloads (default `maxElements=32768`) and fall back to CPU if the difference exceeds a strict tolerance (`abs=1e-9`, `rel=1e-12`). Configure via the `validation` option.
- In `"excel"` mode, if a WebGPU kernel throws at runtime (pipeline/dispatch/device issues), the engine records diagnostics and **falls back to CPU**.
- Kernel edge cases:
  - **Sort** matches `TypedArray#sort` semantics for special values: `NaN` sorts to the end, and `±Infinity` sorts normally.
  - **Histogram** ignores `NaN` values and clamps `±Infinity` into the first/last bin.

## Benchmarks
```bash
pnpm bench:gpu-kernels
```

## Diagnostics UI (example)
Serve the repo with any static server (required for module imports), then open:
`packages/gpu-kernels/examples/diagnostics.html`
