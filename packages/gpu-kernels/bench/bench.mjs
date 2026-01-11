import { performance } from "node:perf_hooks";

import { createKernelEngine } from "../src/index.js";

function formatMs(ms) {
  return `${ms.toFixed(2)}ms`;
}

function formatSpeedup(baselineMs, ms) {
  if (!(ms > 0)) return "n/a";
  return `${(baselineMs / ms).toFixed(2)}x`;
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

  const engineExcel = await createKernelEngine({ precision: "excel", gpu: { enabled: true } });
  const engineFast = await createKernelEngine({ precision: "fast", gpu: { enabled: true } });
  const engineCpu = await createKernelEngine({ gpu: { enabled: false } });

  console.log("Diagnostics (excel):");
  console.log(JSON.stringify(engineExcel.diagnostics(), null, 2));
  console.log("Diagnostics (fast):");
  console.log(JSON.stringify(engineFast.diagnostics(), null, 2));

  const cpuRun = await time("CPU sum", async () => engineCpu.sum(values));
  const excelRun = await time("Excel-mode sum", async () => engineExcel.sum(values));
  const fastRun = await time("Fast-mode sum", async () => engineFast.sum(values));

  console.log(`${cpuRun.name}:  ${formatMs(cpuRun.ms)} result=${cpuRun.result}`);
  console.log(
    `${excelRun.name}: ${formatMs(excelRun.ms)} speedup=${formatSpeedup(cpuRun.ms, excelRun.ms)} backend=${engineExcel.diagnostics().lastKernelBackend.sum} precision=${
      engineExcel.diagnostics().lastKernelPrecision.sum
    }`
  );
  console.log(
    `${fastRun.name}:  ${formatMs(fastRun.ms)} speedup=${formatSpeedup(cpuRun.ms, fastRun.ms)} backend=${engineFast.diagnostics().lastKernelBackend.sum} precision=${
      engineFast.diagnostics().lastKernelPrecision.sum
    }`
  );

  await engineExcel.dispose();
  await engineFast.dispose();
  await engineCpu.dispose();
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
