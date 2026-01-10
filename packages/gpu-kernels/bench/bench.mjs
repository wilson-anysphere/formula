import { performance } from "node:perf_hooks";

import { createKernelEngine } from "../src/index.js";

function formatMs(ms) {
  return `${ms.toFixed(2)}ms`;
}

async function time(name, fn) {
  const t0 = performance.now();
  const result = await fn();
  const t1 = performance.now();
  return { name, ms: t1 - t0, result };
}

async function main() {
  const n = 1_000_000;
  const values = new Float64Array(n);
  for (let i = 0; i < n; i++) values[i] = (i % 1024) * 0.25;

  const engineGpu = await createKernelEngine({ gpu: { enabled: true } });
  const engineCpu = await createKernelEngine({ gpu: { enabled: false } });

  console.log("Diagnostics (GPU-enabled):");
  console.log(JSON.stringify(engineGpu.diagnostics(), null, 2));

  const cpuRun = await time("CPU sum", async () => engineCpu.sum(values));
  const gpuRun = await time("GPU sum (auto)", async () => engineGpu.sum(values));

  console.log(`${cpuRun.name}: ${formatMs(cpuRun.ms)} result=${cpuRun.result}`);
  console.log(`${gpuRun.name}: ${formatMs(gpuRun.ms)} result=${gpuRun.result}`);

  await engineGpu.dispose();
  await engineCpu.dispose();
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});

