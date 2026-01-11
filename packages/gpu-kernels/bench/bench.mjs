import { performance } from "node:perf_hooks";

import { createKernelEngine } from "../src/index.js";

function formatMs(ms) {
  return `${ms.toFixed(2)}ms`;
}

function formatSpeedup(baselineMs, ms) {
  if (!(ms > 0)) return "n/a";
  return `${(baselineMs / ms).toFixed(2)}x`;
}

function formatResult(result) {
  if (result && typeof result === "object" && "length" in result) {
    return `len=${result.length}`;
  }
  if (result && typeof result === "object" && "uniqueKeys" in result) {
    // Group-by kernels
    // @ts-ignore
    return `groups=${result.uniqueKeys?.length ?? 0}`;
  }
  if (result && typeof result === "object" && "leftIndex" in result) {
    // Hash join
    // @ts-ignore
    return `pairs=${result.leftIndex?.length ?? 0}`;
  }
  return String(result);
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
  const values2 = new Float64Array(n);
  for (let i = 0; i < n; i++) values2[i] = (i % 2048) * 0.5;
  const keys = new Uint32Array(n);
  for (let i = 0; i < n; i++) keys[i] = i % 65_536;
  const joinLeftKeys = new Uint32Array(n);
  const joinRightKeys = new Uint32Array(n);
  for (let i = 0; i < n; i++) {
    joinLeftKeys[i] = i;
    joinRightKeys[i] = i;
  }

  const engineExcel = await createKernelEngine({ precision: "excel", gpu: { enabled: true } });
  const engineFast = await createKernelEngine({ precision: "fast", gpu: { enabled: true } });
  const engineCpu = await createKernelEngine({ gpu: { enabled: false } });

  console.log("Diagnostics (excel):");
  console.log(JSON.stringify(engineExcel.diagnostics(), null, 2));
  console.log("Diagnostics (fast):");
  console.log(JSON.stringify(engineFast.diagnostics(), null, 2));

  async function benchKernel(kernel, cpuFn, excelFn, fastFn) {
    const cpuRun = await time(`CPU ${kernel}`, cpuFn);
    const excelRun = await time(`Excel-mode ${kernel}`, excelFn);
    const fastRun = await time(`Fast-mode ${kernel}`, fastFn);

    console.log(`${cpuRun.name}:       ${formatMs(cpuRun.ms)} result=${formatResult(cpuRun.result)}`);
    console.log(
      `${excelRun.name}: ${formatMs(excelRun.ms)} speedup=${formatSpeedup(cpuRun.ms, excelRun.ms)} backend=${
        engineExcel.diagnostics().lastKernelBackend[kernel]
      } precision=${engineExcel.diagnostics().lastKernelPrecision[kernel]}`
    );
    console.log(
      `${fastRun.name}:  ${formatMs(fastRun.ms)} speedup=${formatSpeedup(cpuRun.ms, fastRun.ms)} backend=${
        engineFast.diagnostics().lastKernelBackend[kernel]
      } precision=${engineFast.diagnostics().lastKernelPrecision[kernel]}`
    );
    console.log("");
  }

  await benchKernel(
    "sum",
    async () => engineCpu.sum(values),
    async () => engineExcel.sum(values),
    async () => engineFast.sum(values)
  );

  await benchKernel(
    "sumproduct",
    async () => engineCpu.sumproduct(values, values2),
    async () => engineExcel.sumproduct(values, values2),
    async () => engineFast.sumproduct(values, values2)
  );

  await benchKernel(
    "histogram",
    async () => engineCpu.histogram(values, { min: 0, max: 256, bins: 64 }),
    async () => engineExcel.histogram(values, { min: 0, max: 256, bins: 64 }),
    async () => engineFast.histogram(values, { min: 0, max: 256, bins: 64 })
  );

  await benchKernel(
    "groupBySum",
    async () => engineCpu.groupBySum(keys, values),
    async () => engineExcel.groupBySum(keys, values),
    async () => engineFast.groupBySum(keys, values)
  );

  await benchKernel(
    "hashJoin",
    async () => engineCpu.hashJoin(joinLeftKeys, joinRightKeys),
    async () => engineExcel.hashJoin(joinLeftKeys, joinRightKeys),
    async () => engineFast.hashJoin(joinLeftKeys, joinRightKeys)
  );

  await engineExcel.dispose();
  await engineFast.dispose();
  await engineCpu.dispose();
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
