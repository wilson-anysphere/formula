import test from "node:test";
import assert from "node:assert/strict";

import { KernelEngine } from "../src/index.js";

class FakeGpuBackend {
  kind = "webgpu";
  calls = { sum: 0, sumproduct: 0, mmult: 0, sort: 0, histogram: 0 };
  diagnostics() {
    return { kind: "webgpu", supportedKernels: { sum: true, sumproduct: true, mmult: true, sort: true, histogram: true } };
  }
  dispose() {}
  async sum() {
    this.calls.sum += 1;
    return 123;
  }
  async sumproduct() {
    this.calls.sumproduct += 1;
    return 456;
  }
  async mmult() {
    this.calls.mmult += 1;
    return new Float64Array([0]);
  }
  async sort() {
    this.calls.sort += 1;
    return new Float64Array([0]);
  }
  async histogram() {
    this.calls.histogram += 1;
    return new Uint32Array([0]);
  }
}

test("engine selects CPU for small workloads, GPU for large workloads", async () => {
  const fakeGpu = new FakeGpuBackend();
  const engine = new KernelEngine({
    precision: "fast",
    gpuBackend: fakeGpu,
    gpu: { enabled: true },
    thresholds: { sum: 10 }
  });

  const small = new Float64Array(5).fill(1);
  const large = new Float64Array(100).fill(1);

  await engine.sum(small);
  await engine.sum(large);

  assert.equal(fakeGpu.calls.sum, 1);
  const last = engine.lastKernelBackend();
  assert.equal(last.sum, "webgpu");
});
