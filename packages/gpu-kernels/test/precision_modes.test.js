import test from "node:test";
import assert from "node:assert/strict";

import { KernelEngine } from "../src/index.js";

class FakeGpuBackend {
  kind = "webgpu";
  calls = { sum: 0, min: 0 };
  /** @type {any} */
  lastOpts = null;

  /**
   * @param {boolean} supportsF64
   */
  constructor(supportsF64) {
    this._supportsF64 = supportsF64;
  }

  diagnostics() {
    return { kind: "webgpu", supportedKernels: { sum: true }, supportsF64: this._supportsF64, numericPrecision: "f32" };
  }

  supportsKernelPrecision(kernel, precision) {
    if (precision === "f32") return true;
    return kernel === "sum" || kernel === "min" ? this._supportsF64 : false;
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
