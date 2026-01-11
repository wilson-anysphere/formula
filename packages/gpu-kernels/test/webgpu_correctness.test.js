import test from "node:test";
import assert from "node:assert/strict";

import { WebGpuBackend } from "../src/index.js";
import { CpuBackend } from "../src/index.js";

function approxEqual(a, b, { rel, abs }) {
  const diff = Math.abs(a - b);
  if (diff <= abs) return true;
  return diff <= rel * Math.max(Math.abs(a), Math.abs(b));
}

function makeRng(seed) {
  let state = seed >>> 0;
  return () => {
    state = (1664525 * state + 1013904223) >>> 0;
    return state / 0x1_0000_0000;
  };
}

test("webgpu: f32 + f64 correctness for common kernels (if WebGPU available)", async (t) => {
  const gpu = await WebGpuBackend.createIfSupported();
  if (!gpu) return t.skip("WebGPU unavailable");

  const cpu = new CpuBackend();
  const diag = gpu.diagnostics();

  // -------- f32 paths (explicit) --------
  {
    const rng = makeRng(123);
    const n = 200_000;
    const values = new Float32Array(n);
    for (let i = 0; i < n; i++) values[i] = (rng() - 0.5) * 10_000;

    const cpuSum = await cpu.sum(values);
    const gpuSum = await gpu.sum(values, { precision: "f32" });
    assert.ok(approxEqual(gpuSum, cpuSum, { rel: 1e-4, abs: 1e-3 }), `gpu=${gpuSum} cpu=${cpuSum}`);
  }

  {
    const rng = makeRng(456);
    const n = 200_000;
    const a = new Float32Array(n);
    const b = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      a[i] = (rng() - 0.5) * 100;
      b[i] = (rng() - 0.5) * 100;
    }
    const cpuDot = await cpu.sumproduct(a, b);
    const gpuDot = await gpu.sumproduct(a, b, { precision: "f32" });
    assert.ok(approxEqual(gpuDot, cpuDot, { rel: 1e-3, abs: 1e-2 }), `gpu=${gpuDot} cpu=${cpuDot}`);
  }

  {
    const rng = makeRng(789);
    const n = 100_000;
    const values = new Float32Array(n);
    for (let i = 0; i < n; i++) values[i] = rng();
    values[0] = Number.NaN;
    values[1] = Number.POSITIVE_INFINITY;
    values[2] = Number.NEGATIVE_INFINITY;
    const cpuBins = await cpu.histogram(values, { min: 0, max: 1, bins: 16 });
    const gpuBins = await gpu.histogram(values, { min: 0, max: 1, bins: 16 }, { precision: "f32" });
    assert.deepEqual(Array.from(gpuBins), Array.from(cpuBins));
  }

  {
    const values = new Float32Array([-100, 0, 0.5, 1.0, 2.0]);
    const cpuBins = await cpu.histogram(values, { min: Number.NEGATIVE_INFINITY, max: 1, bins: 2 });
    const gpuBins = await gpu.histogram(values, { min: Number.NEGATIVE_INFINITY, max: 1, bins: 2 }, { precision: "f32" });
    assert.deepEqual(Array.from(gpuBins), Array.from(cpuBins));
  }

  {
    const rng = makeRng(246);
    const n = 200_000;
    const values = new Float32Array(n);
    for (let i = 0; i < n; i++) values[i] = (rng() - 0.5) * 1000;

    const cpuMin = await cpu.min(values);
    const cpuMax = await cpu.max(values);
    const cpuAvg = await cpu.average(values);

    const gpuMin = await gpu.min(values, { precision: "f32" });
    const gpuMax = await gpu.max(values, { precision: "f32" });
    const gpuAvg = await gpu.average(values, { precision: "f32" });

    assert.ok(approxEqual(gpuMin, cpuMin, { rel: 1e-4, abs: 1e-3 }), `gpu=${gpuMin} cpu=${cpuMin}`);
    assert.ok(approxEqual(gpuMax, cpuMax, { rel: 1e-4, abs: 1e-3 }), `gpu=${gpuMax} cpu=${cpuMax}`);
    assert.ok(approxEqual(gpuAvg, cpuAvg, { rel: 1e-4, abs: 1e-3 }), `gpu=${gpuAvg} cpu=${cpuAvg}`);
  }

  {
    const values = new Float32Array([0, -0]);
    const cpuMin = await cpu.min(values);
    const cpuMax = await cpu.max(values);
    const cpuAvg = await cpu.average(values);

    const gpuMin = await gpu.min(values, { precision: "f32" });
    const gpuMax = await gpu.max(values, { precision: "f32" });
    const gpuAvg = await gpu.average(values, { precision: "f32" });

    assert.ok(Object.is(gpuMin, cpuMin), `min: gpu=${gpuMin} cpu=${cpuMin}`);
    assert.ok(Object.is(gpuMax, cpuMax), `max: gpu=${gpuMax} cpu=${cpuMax}`);
    assert.ok(Object.is(gpuAvg, cpuAvg), `avg: gpu=${gpuAvg} cpu=${cpuAvg}`);
  }

  {
    const rng = makeRng(13579);
    const aRows = 32;
    const aCols = 32;
    const bCols = 32;
    const a = new Float32Array(aRows * aCols);
    const b = new Float32Array(aCols * bCols);
    for (let i = 0; i < a.length; i++) a[i] = (rng() - 0.5) * 2;
    for (let i = 0; i < b.length; i++) b[i] = (rng() - 0.5) * 2;

    const cpuOut = await cpu.mmult(a, b, aRows, aCols, bCols);
    const gpuOut = await gpu.mmult(a, b, aRows, aCols, bCols, { precision: "f32" });
    assert.equal(gpuOut.length, cpuOut.length);
    for (let i = 0; i < cpuOut.length; i++) {
      assert.ok(approxEqual(gpuOut[i], cpuOut[i], { rel: 1e-3, abs: 1e-2 }), `i=${i} gpu=${gpuOut[i]} cpu=${cpuOut[i]}`);
    }
  }

  {
    const rng = makeRng(24680);
    const n = 4096;
    const values = new Float32Array(n);
    for (let i = 0; i < n; i++) values[i] = (rng() - 0.5) * 1000;
    values[0] = Number.NaN;
    values[1] = Number.POSITIVE_INFINITY;
    values[2] = 0;
    values[3] = -0;
    const cpuSorted = await cpu.sort(values);
    const gpuSorted = await gpu.sort(values, { precision: "f32" });
    assert.deepEqual(Array.from(gpuSorted), Array.from(cpuSorted));
  }

  {
    const rng = makeRng(4242);
    const n = 50_000;
    const keys = new Uint32Array(n);
    const vals = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      keys[i] = (rng() * 1024) | 0;
      vals[i] = 1;
    }
    const cpuOut = await cpu.groupBySum(keys, vals);
    const gpuOut = await gpu.groupBySum(keys, vals, { precision: "f32" });
    assert.deepEqual(Array.from(gpuOut.uniqueKeys), Array.from(cpuOut.uniqueKeys));
    assert.deepEqual(Array.from(gpuOut.counts), Array.from(cpuOut.counts));
    assert.deepEqual(Array.from(gpuOut.sums), Array.from(cpuOut.sums));
  }

  {
    const rng = makeRng(9999);
    const leftLen = 4096;
    const rightLen = 4096;
    const range = 512;
    const leftKeys = new Uint32Array(leftLen);
    const rightKeys = new Uint32Array(rightLen);
    for (let i = 0; i < leftLen; i++) leftKeys[i] = (rng() * range) | 0;
    for (let i = 0; i < rightLen; i++) rightKeys[i] = (rng() * range) | 0;

    const cpuOut = await cpu.hashJoin(leftKeys, rightKeys);
    const gpuOut = await gpu.hashJoin(leftKeys, rightKeys);
    assert.deepEqual(Array.from(gpuOut.leftIndex), Array.from(cpuOut.leftIndex));
    assert.deepEqual(Array.from(gpuOut.rightIndex), Array.from(cpuOut.rightIndex));
  }

  // -------- f64 paths (when supported) --------
  if (diag.supportsF64) {
    {
      const rng = makeRng(321);
      const n = 200_000;
      const values = new Float64Array(n);
      // Use integer values so the f64 reduction order cannot introduce
      // rounding differences (sum stays < 2^53 and is exactly representable).
      for (let i = 0; i < n; i++) values[i] = ((rng() * 2048) | 0) - 1024;

      const cpuSum = await cpu.sum(values);
      const gpuSum = await gpu.sum(values, { precision: "f64", allowFp32FallbackForF64: false });
      assert.equal(gpuSum, cpuSum);
    }

    {
      const rng = makeRng(654);
      const n = 200_000;
      const a = new Float64Array(n);
      const b = new Float64Array(n);
      for (let i = 0; i < n; i++) {
        // Integer values so products + sums remain exactly representable in f64.
        a[i] = (rng() * 1024) | 0;
        b[i] = (rng() * 1024) | 0;
      }
      const cpuDot = await cpu.sumproduct(a, b);
      const gpuDot = await gpu.sumproduct(a, b, { precision: "f64", allowFp32FallbackForF64: false });
      assert.equal(gpuDot, cpuDot);
    }

    {
      const rng = makeRng(987);
      const n = 100_000;
      const values = new Float64Array(n);
      for (let i = 0; i < n; i++) values[i] = rng();
      values[0] = Number.NaN;
      values[1] = Number.POSITIVE_INFINITY;
      values[2] = Number.NEGATIVE_INFINITY;
      const cpuBins = await cpu.histogram(values, { min: 0, max: 1, bins: 16 });
      const gpuBins = await gpu.histogram(values, { min: 0, max: 1, bins: 16 }, { precision: "f64", allowFp32FallbackForF64: false });
      assert.deepEqual(Array.from(gpuBins), Array.from(cpuBins));
    }

    {
      const values = new Float64Array([-100, 0, 0.5, 1.0, 2.0]);
      const cpuBins = await cpu.histogram(values, { min: Number.NEGATIVE_INFINITY, max: 1, bins: 2 });
      const gpuBins = await gpu.histogram(values, { min: Number.NEGATIVE_INFINITY, max: 1, bins: 2 }, { precision: "f64", allowFp32FallbackForF64: false });
      assert.deepEqual(Array.from(gpuBins), Array.from(cpuBins));
    }

    {
      const rng = makeRng(111);
      const n = 200_000;
      const values = new Float64Array(n);
      for (let i = 0; i < n; i++) values[i] = ((rng() * 2048) | 0) - 1024;

      const cpuMin = await cpu.min(values);
      const cpuMax = await cpu.max(values);
      const cpuAvg = await cpu.average(values);

      const gpuMin = await gpu.min(values, { precision: "f64", allowFp32FallbackForF64: false });
      const gpuMax = await gpu.max(values, { precision: "f64", allowFp32FallbackForF64: false });
      const gpuAvg = await gpu.average(values, { precision: "f64", allowFp32FallbackForF64: false });

      assert.equal(gpuMin, cpuMin);
      assert.equal(gpuMax, cpuMax);
      assert.equal(gpuAvg, cpuAvg);
    }

    {
      const values = new Float64Array([0, -0]);
      const cpuMin = await cpu.min(values);
      const cpuMax = await cpu.max(values);
      const cpuAvg = await cpu.average(values);

      const gpuMin = await gpu.min(values, { precision: "f64", allowFp32FallbackForF64: false });
      const gpuMax = await gpu.max(values, { precision: "f64", allowFp32FallbackForF64: false });
      const gpuAvg = await gpu.average(values, { precision: "f64", allowFp32FallbackForF64: false });

      assert.ok(Object.is(gpuMin, cpuMin), `min: gpu=${gpuMin} cpu=${cpuMin}`);
      assert.ok(Object.is(gpuMax, cpuMax), `max: gpu=${gpuMax} cpu=${cpuMax}`);
      assert.ok(Object.is(gpuAvg, cpuAvg), `avg: gpu=${gpuAvg} cpu=${cpuAvg}`);
    }

    {
      const rng = makeRng(97531);
      const aRows = 32;
      const aCols = 32;
      const bCols = 32;
      const a = new Float64Array(aRows * aCols);
      const b = new Float64Array(aCols * bCols);
      for (let i = 0; i < a.length; i++) a[i] = (rng() - 0.5) * 2;
      for (let i = 0; i < b.length; i++) b[i] = (rng() - 0.5) * 2;

      const cpuOut = await cpu.mmult(a, b, aRows, aCols, bCols);
      const gpuOut = await gpu.mmult(a, b, aRows, aCols, bCols, { precision: "f64", allowFp32FallbackForF64: false });
      assert.equal(gpuOut.length, cpuOut.length);
      for (let i = 0; i < cpuOut.length; i++) {
        assert.ok(approxEqual(gpuOut[i], cpuOut[i], { rel: 1e-12, abs: 1e-9 }), `i=${i} gpu=${gpuOut[i]} cpu=${cpuOut[i]}`);
      }
    }

    {
      const rng = makeRng(86420);
      const n = 4096;
      const values = new Float64Array(n);
      for (let i = 0; i < n; i++) values[i] = (rng() - 0.5) * 1000;
      values[0] = Number.NaN;
      values[1] = Number.POSITIVE_INFINITY;
      values[2] = 0;
      values[3] = -0;
      const cpuSorted = await cpu.sort(values);
      const gpuSorted = await gpu.sort(values, { precision: "f64", allowFp32FallbackForF64: false });
      assert.deepEqual(Array.from(gpuSorted), Array.from(cpuSorted));
    }
  }

  gpu.dispose();
});
