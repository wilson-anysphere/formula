import test from "node:test";
import assert from "node:assert/strict";

import { WebGpuBackend } from "../src/index.js";

function approxEqual(a, b, { rel = 1e-4, abs = 1e-3 } = {}) {
  const diff = Math.abs(a - b);
  if (diff <= abs) return true;
  return diff <= rel * Math.max(Math.abs(a), Math.abs(b));
}

test("webgpu: compares SUM against CPU within tolerance (if WebGPU available)", async (t) => {
  const gpu = await WebGpuBackend.createIfSupported();
  if (!gpu) return t.skip("WebGPU unavailable");

  const n = 200_000;
  const values = new Float64Array(n);
  for (let i = 0; i < n; i++) values[i] = (i % 1024) * 0.25;

  let cpu = 0;
  for (let i = 0; i < n; i++) cpu += values[i];

  const gpuSum = await gpu.sum(values, { allowFp32FallbackForF64: true });
  assert.ok(approxEqual(gpuSum, cpu), `gpu=${gpuSum} cpu=${cpu}`);
  gpu.dispose();
});

test("webgpu: compares SUMPRODUCT against CPU within tolerance (if WebGPU available)", async (t) => {
  const gpu = await WebGpuBackend.createIfSupported();
  if (!gpu) return t.skip("WebGPU unavailable");

  const n = 200_000;
  const a = new Float64Array(n);
  const b = new Float64Array(n);
  for (let i = 0; i < n; i++) {
    a[i] = (i % 1024) * 0.1;
    b[i] = (i % 2048) * 0.2;
  }

  let cpu = 0;
  for (let i = 0; i < n; i++) cpu += a[i] * b[i];

  const gpuDot = await gpu.sumproduct(a, b, { allowFp32FallbackForF64: true });
  assert.ok(approxEqual(gpuDot, cpu), `gpu=${gpuDot} cpu=${cpu}`);
  gpu.dispose();
});

