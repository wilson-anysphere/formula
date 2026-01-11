import test from "node:test";
import assert from "node:assert/strict";

import { KernelEngine } from "../src/index.js";

class FakeGpuBackend {
  kind = "webgpu";
  calls = { sum: 0, min: 0, sort: 0, histogram: 0 };
  /** @type {any} */
  lastOpts = null;

  /**
   * @param {boolean} supportsF64
   */
  constructor(supportsF64) {
    this._supportsF64 = supportsF64;
  }

  diagnostics() {
    return {
      kind: "webgpu",
      supportedKernels: { sum: true, min: true, sort: true, histogram: true },
      supportsF64: this._supportsF64,
      numericPrecision: "f32"
    };
  }

  supportsKernelPrecision(kernel, precision) {
    if (precision === "f32") return true;
    return kernel === "sum" || kernel === "min" || kernel === "sort" || kernel === "histogram" ? this._supportsF64 : false;
  }

  dispose() {}

  async sum(_values, opts) {
    this.calls.sum += 1;
    this.lastOpts = opts;
    return 123;
  }

  async min(_values, opts) {
    this.calls.min += 1;
    this.lastOpts = opts;
    return 5;
  }

  async sort(values, opts) {
    this.calls.sort += 1;
    this.lastOpts = opts;
    const out = Float64Array.from(values);
    out.sort();
    return out;
  }

  async histogram(values, opts, backendOpts) {
    this.calls.histogram += 1;
    this.lastOpts = backendOpts;
    const { min, max, bins } = opts;
    if (!(bins > 0)) throw new Error("histogram bins must be > 0");
    if (!(max > min)) throw new Error("histogram max must be > min");
    const counts = new Uint32Array(bins);
    const invWidth = bins / (max - min);
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      if (Number.isNaN(v)) continue;
      let bin = Math.floor((v - min) * invWidth);
      if (bin < 0) bin = 0;
      if (bin >= bins) bin = bins - 1;
      counts[bin] += 1;
    }
    return counts;
  }
}

test("precision=excel: falls back to CPU when GPU cannot do f64", async () => {
  const fakeGpu = new FakeGpuBackend(false);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 }
  });

  const values = new Float64Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 6);
  assert.equal(fakeGpu.calls.sum, 0);
  assert.equal(engine.lastKernelBackend().sum, "cpu");
});

test("precision=excel: also falls back to CPU for Float32Array inputs when GPU cannot do f64", async () => {
  const fakeGpu = new FakeGpuBackend(false);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 }
  });

  const values = new Float32Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 6);
  assert.equal(fakeGpu.calls.sum, 0);
  assert.equal(engine.lastKernelBackend().sum, "cpu");
});

test("precision=excel: forces f64 precision on GPU when available", async () => {
  const fakeGpu = new FakeGpuBackend(true);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 },
    validation: { enabled: false }
  });

  const values = new Float32Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 123);
  assert.equal(fakeGpu.calls.sum, 1);
  assert.equal(engine.lastKernelBackend().sum, "webgpu");

  const diag = engine.diagnostics();
  assert.equal(diag.lastKernelPrecision.sum, "f64");
  assert.deepEqual(fakeGpu.lastOpts, { precision: "f64", allowFp32FallbackForF64: false });
});

test("precision=excel: validation falls back to CPU on mismatch and records diagnostics", async () => {
  class WrongGpuBackend extends FakeGpuBackend {
    async sum() {
      this.calls.sum += 1;
      return 0;
    }
  }
  const fakeGpu = new WrongGpuBackend(true);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 }
  });

  const values = new Float64Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 6);
  assert.equal(engine.lastKernelBackend().sum, "cpu");

  const diag = engine.diagnostics();
  assert.equal(diag.validation.mismatches, 1);
  assert.equal(diag.validation.lastMismatch.kernel, "sum");
});

test("excel mode: falls back to CPU on GPU error and records diagnostics", async () => {
  class ErrorGpuBackend extends FakeGpuBackend {
    async sum() {
      this.calls.sum += 1;
      throw new Error("GPU failed");
    }
  }

  const fakeGpu = new ErrorGpuBackend(true);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 },
    validation: { enabled: false }
  });

  const values = new Float64Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 6);
  assert.equal(engine.lastKernelBackend().sum, "cpu");

  const diag = engine.diagnostics();
  assert.equal(diag.validation.gpuErrors, 1);
  assert.equal(diag.validation.lastGpuError.kernel, "sum");
});

test("excel mode: validation treats +0 vs -0 mismatches as mismatches", async () => {
  class ZeroMinGpuBackend extends FakeGpuBackend {
    async min() {
      this.calls.min += 1;
      return 0;
    }
  }

  const fakeGpu = new ZeroMinGpuBackend(true);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { min: 0 }
  });

  const values = new Float64Array([0, -0]);
  const result = await engine.min(values);
  assert.ok(Object.is(result, -0), `expected -0, got ${result}`);

  const diag = engine.diagnostics();
  assert.equal(diag.validation.mismatches, 1);
  assert.equal(diag.validation.lastMismatch.kernel, "min");
});

test("precision=fast: uses f32 on GPU and does not validate by default", async () => {
  class WrongGpuBackend extends FakeGpuBackend {
    async sum() {
      this.calls.sum += 1;
      return 0;
    }
  }

  const fakeGpu = new WrongGpuBackend(true);
  const engine = new KernelEngine({
    precision: "fast",
    gpuBackend: fakeGpu,
    thresholds: { sum: 0 }
  });

  const values = new Float64Array([1, 2, 3]);
  const result = await engine.sum(values);
  assert.equal(result, 0);
  assert.equal(engine.lastKernelBackend().sum, "webgpu");

  const diag = engine.diagnostics();
  assert.equal(diag.validation.enabled, false);
  assert.equal(diag.lastKernelPrecision.sum, "f32");
});

test("new kernels: min follows the same excel/fast precision rules", async () => {
  {
    const fakeGpu = new FakeGpuBackend(false);
    const engine = new KernelEngine({
      precision: "excel",
      gpuBackend: fakeGpu,
      thresholds: { min: 0 }
    });

    const values = new Float64Array([3, 1, 2]);
    const result = await engine.min(values);
    assert.equal(result, 1);
    assert.equal(fakeGpu.calls.min, 0);
    assert.equal(engine.lastKernelBackend().min, "cpu");
  }

  {
    const fakeGpu = new FakeGpuBackend(true);
    const engine = new KernelEngine({
      precision: "excel",
      gpuBackend: fakeGpu,
      thresholds: { min: 0 },
      validation: { enabled: false }
    });

    const values = new Float32Array([3, 1, 2]);
    const result = await engine.min(values);
    assert.equal(result, 5);
    assert.equal(fakeGpu.calls.min, 1);
    assert.equal(engine.lastKernelBackend().min, "webgpu");
    assert.equal(engine.diagnostics().lastKernelPrecision.min, "f64");
    assert.deepEqual(fakeGpu.lastOpts, { precision: "f64", allowFp32FallbackForF64: false });
  }

  {
    const fakeGpu = new FakeGpuBackend(true);
    const engine = new KernelEngine({
      precision: "fast",
      gpuBackend: fakeGpu,
      thresholds: { min: 0 }
    });

    const values = new Float64Array([3, 1, 2]);
    const result = await engine.min(values);
    assert.equal(result, 5);
    assert.equal(engine.lastKernelBackend().min, "webgpu");
    assert.equal(engine.diagnostics().lastKernelPrecision.min, "f32");
  }
});

test("count: always uses CPU and is exact", async () => {
  const fakeGpu = new FakeGpuBackend(true);
  const engine = new KernelEngine({
    precision: "fast",
    gpuBackend: fakeGpu,
    thresholds: { count: 0 }
  });

  const values = new Float64Array(123).fill(0);
  const result = await engine.count(values);
  assert.equal(result, 123);
  assert.equal(engine.lastKernelBackend().count, "cpu");
  assert.equal(engine.diagnostics().lastKernelPrecision.count, "f64");
});

test("sort: never downcasts Float64Array even in fast mode", async () => {
  const fakeGpu = new FakeGpuBackend(false);
  const engine = new KernelEngine({
    precision: "fast",
    gpuBackend: fakeGpu,
    thresholds: { sort: 0 }
  });

  const values = new Float64Array([3, 1, 2]);
  const out = await engine.sort(values);
  assert.deepEqual(Array.from(out), [1, 2, 3]);
  assert.equal(fakeGpu.calls.sort, 0);
  assert.equal(engine.lastKernelBackend().sort, "cpu");
  assert.equal(engine.diagnostics().lastKernelPrecision.sort, "f64");
});

test("histogram: excel mode requires f64 GPU support even for Float32Array inputs", async () => {
  const fakeGpu = new FakeGpuBackend(false);
  const engine = new KernelEngine({
    precision: "excel",
    gpuBackend: fakeGpu,
    thresholds: { histogram: 0 }
  });

  const values = new Float32Array([0.1, 0.2, 0.9]);
  const out = await engine.histogram(values, { min: 0, max: 1, bins: 2 });
  assert.deepEqual(Array.from(out), [2, 1]);
  assert.equal(fakeGpu.calls.histogram, 0);
  assert.equal(engine.lastKernelBackend().histogram, "cpu");
});

test("histogram: fast mode uses f32 precision for Float64Array inputs", async () => {
  const fakeGpu = new FakeGpuBackend(true);
  const engine = new KernelEngine({
    precision: "fast",
    gpuBackend: fakeGpu,
    thresholds: { histogram: 0 }
  });

  const values = new Float64Array([0.1, 0.2, 0.9]);
  const out = await engine.histogram(values, { min: 0, max: 1, bins: 2 });
  assert.deepEqual(Array.from(out), [2, 1]);
  assert.equal(fakeGpu.calls.histogram, 1);
  assert.deepEqual(fakeGpu.lastOpts, { precision: "f32", allowFp32FallbackForF64: true });
  assert.equal(engine.lastKernelBackend().histogram, "webgpu");
  assert.equal(engine.diagnostics().lastKernelPrecision.histogram, "f32");
});
