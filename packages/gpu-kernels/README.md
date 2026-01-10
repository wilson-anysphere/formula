# @formula/gpu-kernels
Optional WebGPU compute acceleration for heavy spreadsheet/analytics kernels with CPU fallback.

## What’s implemented
- **SUM / SUMPRODUCT** (parallel reduction)
- **MMULT** (matrix multiplication)
- **Sort** (bitonic sort for numeric vectors)
- **Histogram/binning** (atomic bin counts)

All kernels have a **CPU fallback** and will run without WebGPU.

## Usage (high-level engine)
```js
import { createKernelEngine } from "./src/index.js";

const engine = await createKernelEngine({
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
- The current GPU implementation uses **f32** arithmetic internally. The `KernelEngine` compares results within tolerance in tests and is intended for acceleration of large workloads where small floating-point differences are acceptable.

## Benchmarks
```bash
pnpm bench:gpu-kernels
```

## Diagnostics UI (example)
Serve the repo with any static server (required for module imports), then open:
`packages/gpu-kernels/examples/diagnostics.html`
