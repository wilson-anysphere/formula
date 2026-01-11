import { CpuBackend } from "./cpu_backend.js";
import { WebGpuBackend } from "./webgpu_backend.js";

export const DEFAULT_THRESHOLDS = {
  sum: 1 << 15,
  min: 1 << 15,
  max: 1 << 15,
  average: 1 << 15,
  count: 1 << 15,
  sumproduct: 1 << 15,
  groupByCount: 1 << 15,
  groupBySum: 1 << 15,
  groupByMin: 1 << 15,
  groupByMax: 1 << 15,
  groupByCount2: 1 << 15,
  groupBySum2: 1 << 15,
  groupByMin2: 1 << 15,
  groupByMax2: 1 << 15,
  hashJoin: 1 << 15,
  // Rough heuristic based on multiply-add count (aRows * aCols * bCols).
  mmult: 1 << 20,
  sort: 1 << 15,
  histogram: 1 << 15
};

const DEFAULT_VALIDATION = {
  enabled: false,
  maxElements: 1 << 15,
  absTolerance: 1e-9,
  relTolerance: 1e-12
};

/**
 * @typedef {{
 *  precision?: "excel" | "fast",
 *  gpu?: {
 *    enabled?: boolean,
 *    forceBackend?: "auto" | "cpu" | "gpu",
 *    allowFp32FallbackForF64?: boolean
 *  },
 *  validation?: Partial<typeof DEFAULT_VALIDATION>,
 *  thresholds?: Partial<typeof DEFAULT_THRESHOLDS>,
 *  cpuBackend?: CpuBackend,
 *  gpuBackend?: WebGpuBackend | null
 * }} KernelEngineOptions
 */

export class KernelEngine {
  /**
   * @param {KernelEngineOptions} opts
   */
  constructor(opts) {
    this._cpu = opts.cpuBackend ?? new CpuBackend();
    this._gpu = opts.gpuBackend ?? null;
    this._precision = opts.precision ?? "excel";
    this._gpuEnabled = opts.gpu?.enabled ?? true;
    this._forceBackend = opts.gpu?.forceBackend ?? "auto";
    // Excel mode must never silently downcast f64->f32.
    this._allowFp32FallbackForF64 = this._precision === "excel" ? false : (opts.gpu?.allowFp32FallbackForF64 ?? true);
    this._thresholds = { ...DEFAULT_THRESHOLDS, ...(opts.thresholds ?? {}) };
    this._validation = {
      ...DEFAULT_VALIDATION,
      enabled: this._precision === "excel",
      ...(opts.validation ?? {})
    };
    this._validationState = { mismatches: 0, lastMismatch: null, gpuErrors: 0, lastGpuError: null };

    /** @type {Record<string, "cpu" | "webgpu">} */
    this._lastKernelBackend = {
      sum: "cpu",
      min: "cpu",
      max: "cpu",
      average: "cpu",
      count: "cpu",
      sumproduct: "cpu",
      groupByCount: "cpu",
      groupBySum: "cpu",
      groupByMin: "cpu",
      groupByMax: "cpu",
      groupByCount2: "cpu",
      groupBySum2: "cpu",
      groupByMin2: "cpu",
      groupByMax2: "cpu",
      hashJoin: "cpu",
      mmult: "cpu",
      sort: "cpu",
      histogram: "cpu"
    };

    /** @type {Record<string, "f32" | "f64" | "u32">} */
    this._lastKernelPrecision = {
      sum: "f64",
      min: "f64",
      max: "f64",
      average: "f64",
      count: "f64",
      sumproduct: "f64",
      groupByCount: "u32",
      groupBySum: "f64",
      groupByMin: "f64",
      groupByMax: "f64",
      groupByCount2: "u32",
      groupBySum2: "f64",
      groupByMin2: "f64",
      groupByMax2: "f64",
      hashJoin: "u32",
      mmult: "f64",
      sort: "f64",
      histogram: "f64"
    };
  }

  /**
   * @param {KernelEngineOptions} opts
   */
  static async create(opts = {}) {
    const gpuEnabled = opts.gpu?.enabled ?? true;
    const gpuBackend = opts.gpuBackend ?? (gpuEnabled ? await WebGpuBackend.createIfSupported() : null);
    return new KernelEngine({ ...opts, gpuBackend });
  }

  /**
   * @returns {boolean}
   */
  get gpuEnabled() {
    return this._gpuEnabled;
  }

  /**
   * @param {boolean} enabled
   */
  set gpuEnabled(enabled) {
    this._gpuEnabled = enabled;
  }

  /**
   * @returns {import("./cpu_backend.js").CpuBackend}
   */
  get cpuBackend() {
    return this._cpu;
  }

  /**
   * @returns {import("./webgpu_backend.js").WebGpuBackend | null}
   */
  get gpuBackend() {
    return this._gpu;
  }

  /**
   * For UI/telemetry. Updated on each kernel call.
   */
  lastKernelBackend() {
    return { ...this._lastKernelBackend };
  }

  diagnostics() {
    return {
      precision: this._precision,
      gpu: {
        enabled: this._gpuEnabled,
        forceBackend: this._forceBackend,
        allowFp32FallbackForF64: this._allowFp32FallbackForF64,
        available: Boolean(this._gpu),
        ...(this._gpu ? this._gpu.diagnostics() : { kind: "webgpu", supportedKernels: {} })
      },
      cpu: this._cpu.diagnostics(),
      thresholds: { ...this._thresholds },
      lastKernelBackend: this.lastKernelBackend(),
      lastKernelPrecision: { ...this._lastKernelPrecision },
      validation: { ...this._validation, ...this._validationState }
    };
  }

  async dispose() {
    if (this._gpu) this._gpu.dispose();
  }

  /**
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "groupByCount" | "groupBySum" | "groupByMin" | "groupByMax" | "groupByCount2" | "groupBySum2" | "groupByMin2" | "groupByMax2" | "hashJoin" | "mmult" | "sort" | "histogram"} kernel
   * @param {number} workloadSize
   * @param {"f32" | "f64" | "u32"} gpuPrecision
   * @returns {"cpu" | "webgpu"}
   */
  _chooseBackend(kernel, workloadSize, gpuPrecision) {
    if (this._forceBackend === "cpu") return "cpu";
    if (this._forceBackend === "gpu") {
      if (!this._gpu || !this._gpuEnabled) return "cpu";
      return this._gpuSupports(kernel, gpuPrecision) ? "webgpu" : "cpu";
    }
    if (!this._gpu || !this._gpuEnabled) return "cpu";

    const threshold = this._thresholds[kernel] ?? Infinity;
    if (workloadSize < threshold) return "cpu";
    return this._gpuSupports(kernel, gpuPrecision) ? "webgpu" : "cpu";
  }

  /**
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "groupByCount" | "groupBySum" | "groupByMin" | "groupByMax" | "groupByCount2" | "groupBySum2" | "groupByMin2" | "groupByMax2" | "hashJoin" | "mmult" | "sort" | "histogram"} kernel
   * @param {"f32" | "f64" | "u32"} precision
   */
  _gpuSupports(kernel, precision) {
    if (!this._gpu) return false;
    if (precision === "f32") return true;
    if (typeof this._gpu.supportsKernelPrecision === "function") {
      return this._gpu.supportsKernelPrecision(kernel, precision);
    }
    return false;
  }

  /**
   * @param {number} gpu
   * @param {number} cpu
   */
  _withinTolerance(gpu, cpu) {
    if (Object.is(gpu, cpu)) return true;
    // Treat +0 and -0 as distinct: they can affect downstream formulas
    // (e.g. 1/(-0) => -Infinity).
    if (gpu === 0 && cpu === 0) return false;
    if (!Number.isFinite(gpu) || !Number.isFinite(cpu)) return false;
    const diff = Math.abs(gpu - cpu);
    if (diff <= this._validation.absTolerance) return true;
    return diff <= this._validation.relTolerance * Math.max(Math.abs(gpu), Math.abs(cpu));
  }

  /**
   * @param {"sum" | "min" | "max" | "average" | "sumproduct"} kernel
   * @param {number} gpu
   * @param {number} cpu
   * @param {number} workloadSize
   * @param {"f32" | "f64"} precision
   */
  _recordValidationMismatch(kernel, gpu, cpu, workloadSize, precision) {
    this._validationState.mismatches += 1;
    this._validationState.lastMismatch = {
      kernel,
      precision,
      workloadSize,
      gpu,
      cpu,
      absDiff: Math.abs(gpu - cpu)
    };
  }

  /**
   * @param {string} kernel
   * @param {"f32" | "f64" | "u32"} precision
   * @param {unknown} err
   */
  _recordGpuError(kernel, precision, err) {
    this._validationState.gpuErrors += 1;
    this._validationState.lastGpuError = {
      kernel,
      precision,
      message: err instanceof Error ? err.message : String(err)
    };
  }

  /**
   * @param {"sum" | "min" | "max" | "average" | "sumproduct" | "histogram" | "groupByCount" | "groupBySum" | "groupByMin" | "groupByMax" | "groupByCount2" | "groupBySum2" | "groupByMin2" | "groupByMax2" | "hashJoin"} kernel
   * @param {number} workloadSize
   */
  _shouldValidate(kernel, workloadSize) {
    if (!this._validation.enabled) return false;
    if (workloadSize > this._validation.maxElements) return false;
    return (
      kernel === "sum" ||
      kernel === "sumproduct" ||
      kernel === "histogram" ||
      kernel === "min" ||
      kernel === "max" ||
      kernel === "average" ||
      kernel === "groupByCount" ||
      kernel === "groupBySum" ||
      kernel === "groupByMin" ||
      kernel === "groupByMax" ||
      kernel === "groupByCount2" ||
      kernel === "groupBySum2" ||
      kernel === "groupByMin2" ||
      kernel === "groupByMax2" ||
      kernel === "hashJoin"
    );
  }

  /**
   * @param {"f32" | "f64"} requested
   * @param {ArrayLike<number>} values
   */
  _gpuPrecisionForValues(requested, values) {
    if (requested === "f64") return "f64";
    if (this._precision === "excel") return "f64";
    if (!this._allowFp32FallbackForF64 && values instanceof Float64Array) return "f64";
    return "f32";
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async sum(values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = values.length;
    const backend = this._chooseBackend("sum", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.sum(values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("sum", workloadSize)) {
          const cpu = await this._cpu.sum(values);
          if (!this._withinTolerance(gpu, cpu)) {
            this._recordValidationMismatch("sum", gpu, cpu, workloadSize, gpuPrecision);
            this._lastKernelBackend.sum = "cpu";
            this._lastKernelPrecision.sum = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.sum = "webgpu";
        this._lastKernelPrecision.sum = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("sum", gpuPrecision, err);
        this._lastKernelBackend.sum = "cpu";
        this._lastKernelPrecision.sum = "f64";
        return this._cpu.sum(values);
      }
    }

    this._lastKernelBackend.sum = "cpu";
    this._lastKernelPrecision.sum = "f64";
    return this._cpu.sum(values);
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   */
  async sumproduct(a, b) {
    const workloadSize = a.length;
    const anyF64 = a instanceof Float64Array || b instanceof Float64Array;
    const gpuPrecision = this._precision === "excel" ? "f64" : !this._allowFp32FallbackForF64 && anyF64 ? "f64" : "f32";
    const backend = this._chooseBackend("sumproduct", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.sumproduct(a, b, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("sumproduct", workloadSize)) {
          const cpu = await this._cpu.sumproduct(a, b);
          if (!this._withinTolerance(gpu, cpu)) {
            this._recordValidationMismatch("sumproduct", gpu, cpu, workloadSize, gpuPrecision);
            this._lastKernelBackend.sumproduct = "cpu";
            this._lastKernelPrecision.sumproduct = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.sumproduct = "webgpu";
        this._lastKernelPrecision.sumproduct = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("sumproduct", gpuPrecision, err);
        this._lastKernelBackend.sumproduct = "cpu";
        this._lastKernelPrecision.sumproduct = "f64";
        return this._cpu.sumproduct(a, b);
      }
    }

    this._lastKernelBackend.sumproduct = "cpu";
    this._lastKernelPrecision.sumproduct = "f64";
    return this._cpu.sumproduct(a, b);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async min(values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = values.length;
    const backend = this._chooseBackend("min", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.min(values, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
        if (this._shouldValidate("min", workloadSize)) {
          const cpu = await this._cpu.min(values);
          if (!this._withinTolerance(gpu, cpu)) {
            this._recordValidationMismatch("min", gpu, cpu, workloadSize, gpuPrecision);
            this._lastKernelBackend.min = "cpu";
            this._lastKernelPrecision.min = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.min = "webgpu";
        this._lastKernelPrecision.min = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("min", gpuPrecision, err);
        this._lastKernelBackend.min = "cpu";
        this._lastKernelPrecision.min = "f64";
        return this._cpu.min(values);
      }
    }

    this._lastKernelBackend.min = "cpu";
    this._lastKernelPrecision.min = "f64";
    return this._cpu.min(values);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async max(values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = values.length;
    const backend = this._chooseBackend("max", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.max(values, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
        if (this._shouldValidate("max", workloadSize)) {
          const cpu = await this._cpu.max(values);
          if (!this._withinTolerance(gpu, cpu)) {
            this._recordValidationMismatch("max", gpu, cpu, workloadSize, gpuPrecision);
            this._lastKernelBackend.max = "cpu";
            this._lastKernelPrecision.max = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.max = "webgpu";
        this._lastKernelPrecision.max = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("max", gpuPrecision, err);
        this._lastKernelBackend.max = "cpu";
        this._lastKernelPrecision.max = "f64";
        return this._cpu.max(values);
      }
    }

    this._lastKernelBackend.max = "cpu";
    this._lastKernelPrecision.max = "f64";
    return this._cpu.max(values);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async count(values) {
    // Counting a typed array is O(1) and always exact; do it on CPU.
    this._lastKernelBackend.count = "cpu";
    this._lastKernelPrecision.count = "f64";
    return this._cpu.count(values);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async average(values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = values.length;
    const backend = this._chooseBackend("average", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.average(values, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
        if (this._shouldValidate("average", workloadSize)) {
          const cpu = await this._cpu.average(values);
          if (!this._withinTolerance(gpu, cpu)) {
            this._recordValidationMismatch("average", gpu, cpu, workloadSize, gpuPrecision);
            this._lastKernelBackend.average = "cpu";
            this._lastKernelPrecision.average = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.average = "webgpu";
        this._lastKernelPrecision.average = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("average", gpuPrecision, err);
        this._lastKernelBackend.average = "cpu";
        this._lastKernelPrecision.average = "f64";
        return this._cpu.average(values);
      }
    }

    this._lastKernelBackend.average = "cpu";
    this._lastKernelPrecision.average = "f64";
    return this._cpu.average(values);
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   * @param {number} aRows
   * @param {number} aCols
   * @param {number} bCols
   */
  async mmult(a, b, aRows, aCols, bCols) {
    const mulAdds = aRows * aCols * bCols;
    const anyF64 = a instanceof Float64Array || b instanceof Float64Array;
    const gpuPrecision = this._precision === "excel" ? "f64" : !this._allowFp32FallbackForF64 && anyF64 ? "f64" : "f32";
    const backend = this._chooseBackend("mmult", mulAdds, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const out = await this._gpu.mmult(a, b, aRows, aCols, bCols, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        this._lastKernelBackend.mmult = "webgpu";
        this._lastKernelPrecision.mmult = gpuPrecision;
        return out;
      } catch (err) {
        this._recordGpuError("mmult", gpuPrecision, err);
        this._lastKernelBackend.mmult = "cpu";
        this._lastKernelPrecision.mmult = "f64";
        return this._cpu.mmult(a, b, aRows, aCols, bCols);
      }
    }
    this._lastKernelBackend.mmult = "cpu";
    this._lastKernelPrecision.mmult = "f64";
    return this._cpu.mmult(a, b, aRows, aCols, bCols);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async sort(values) {
    const workloadSize = values.length;
    // Sorting is order-sensitive; never silently downcast Float64Array -> f32.
    const gpuPrecision = values instanceof Float64Array ? "f64" : this._gpuPrecisionForValues("f32", values);
    const backend = this._chooseBackend("sort", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const out = await this._gpu.sort(values, { precision: gpuPrecision, allowFp32FallbackForF64: false });
        this._lastKernelBackend.sort = "webgpu";
        this._lastKernelPrecision.sort = gpuPrecision;
        return out;
      } catch (err) {
        this._recordGpuError("sort", gpuPrecision, err);
        this._lastKernelBackend.sort = "cpu";
        this._lastKernelPrecision.sort = "f64";
        return this._cpu.sort(values);
      }
    }
    this._lastKernelBackend.sort = "cpu";
    this._lastKernelPrecision.sort = "f64";
    return this._cpu.sort(values);
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ min: number, max: number, bins: number }} opts
   */
  async histogram(values, opts) {
    const workloadSize = values.length;
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const backend = this._chooseBackend("histogram", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.histogram(values, opts, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("histogram", workloadSize)) {
          const cpu = await this._cpu.histogram(values, opts);
          let ok = cpu.length === gpu.length;
          if (ok) {
            for (let i = 0; i < cpu.length; i++) {
              if (cpu[i] !== gpu[i]) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "histogram", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.histogram = "cpu";
            this._lastKernelPrecision.histogram = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.histogram = "webgpu";
        this._lastKernelPrecision.histogram = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("histogram", gpuPrecision, err);
        this._lastKernelBackend.histogram = "cpu";
        this._lastKernelPrecision.histogram = "f64";
        return this._cpu.histogram(values, opts);
      }
    }

    this._lastKernelBackend.histogram = "cpu";
    this._lastKernelPrecision.histogram = "f64";
    return this._cpu.histogram(values, opts);
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   */
  async groupByCount(keys) {
    const workloadSize = keys.length;
    const backend = this._chooseBackend("groupByCount", workloadSize, "u32");

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.groupByCount(keys);
        if (this._shouldValidate("groupByCount", workloadSize)) {
          const cpu = await this._cpu.groupByCount(keys);
          let ok = cpu.uniqueKeys.length === gpu.uniqueKeys.length && cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.uniqueKeys.length; i++) {
              if (cpu.uniqueKeys[i] !== gpu.uniqueKeys[i] || cpu.counts[i] !== gpu.counts[i]) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByCount", precision: "u32", workloadSize };
            this._lastKernelBackend.groupByCount = "cpu";
            this._lastKernelPrecision.groupByCount = "u32";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByCount = "webgpu";
        this._lastKernelPrecision.groupByCount = "u32";
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByCount", "u32", err);
        this._lastKernelBackend.groupByCount = "cpu";
        this._lastKernelPrecision.groupByCount = "u32";
        return this._cpu.groupByCount(keys);
      }
    }

    this._lastKernelBackend.groupByCount = "cpu";
    this._lastKernelPrecision.groupByCount = "u32";
    return this._cpu.groupByCount(keys);
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   */
  async groupBySum(keys, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keys.length;
    const backend = this._chooseBackend("groupBySum", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.groupBySum(keys, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupBySum", workloadSize)) {
          const cpu = await this._cpu.groupBySum(keys, values);
          let ok =
            cpu.uniqueKeys.length === gpu.uniqueKeys.length &&
            cpu.sums.length === gpu.sums.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.uniqueKeys.length; i++) {
              if (cpu.uniqueKeys[i] !== gpu.uniqueKeys[i] || cpu.counts[i] !== gpu.counts[i] || !this._withinTolerance(gpu.sums[i], cpu.sums[i])) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupBySum", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupBySum = "cpu";
            this._lastKernelPrecision.groupBySum = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupBySum = "webgpu";
        this._lastKernelPrecision.groupBySum = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupBySum", gpuPrecision, err);
        this._lastKernelBackend.groupBySum = "cpu";
        this._lastKernelPrecision.groupBySum = "f64";
        return this._cpu.groupBySum(keys, values);
      }
    }

    this._lastKernelBackend.groupBySum = "cpu";
    this._lastKernelPrecision.groupBySum = "f64";
    return this._cpu.groupBySum(keys, values);
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   */
  async groupByMin(keys, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keys.length;
    const backend = this._chooseBackend("groupByMin", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.groupByMin(keys, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupByMin", workloadSize)) {
          const cpu = await this._cpu.groupByMin(keys, values);
          let ok =
            cpu.uniqueKeys.length === gpu.uniqueKeys.length &&
            cpu.mins.length === gpu.mins.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.uniqueKeys.length; i++) {
              if (cpu.uniqueKeys[i] !== gpu.uniqueKeys[i] || cpu.counts[i] !== gpu.counts[i] || !Object.is(gpu.mins[i], cpu.mins[i])) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByMin", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupByMin = "cpu";
            this._lastKernelPrecision.groupByMin = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByMin = "webgpu";
        this._lastKernelPrecision.groupByMin = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByMin", gpuPrecision, err);
        this._lastKernelBackend.groupByMin = "cpu";
        this._lastKernelPrecision.groupByMin = "f64";
        return this._cpu.groupByMin(keys, values);
      }
    }

    this._lastKernelBackend.groupByMin = "cpu";
    this._lastKernelPrecision.groupByMin = "f64";
    return this._cpu.groupByMin(keys, values);
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   */
  async groupByMax(keys, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keys.length;
    const backend = this._chooseBackend("groupByMax", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.groupByMax(keys, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupByMax", workloadSize)) {
          const cpu = await this._cpu.groupByMax(keys, values);
          let ok =
            cpu.uniqueKeys.length === gpu.uniqueKeys.length &&
            cpu.maxs.length === gpu.maxs.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.uniqueKeys.length; i++) {
              if (cpu.uniqueKeys[i] !== gpu.uniqueKeys[i] || cpu.counts[i] !== gpu.counts[i] || !Object.is(gpu.maxs[i], cpu.maxs[i])) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByMax", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupByMax = "cpu";
            this._lastKernelPrecision.groupByMax = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByMax = "webgpu";
        this._lastKernelPrecision.groupByMax = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByMax", gpuPrecision, err);
        this._lastKernelBackend.groupByMax = "cpu";
        this._lastKernelPrecision.groupByMax = "f64";
        return this._cpu.groupByMax(keys, values);
      }
    }

    this._lastKernelBackend.groupByMax = "cpu";
    this._lastKernelPrecision.groupByMax = "f64";
    return this._cpu.groupByMax(keys, values);
  }

  /**
   * Two-key group-by COUNT. CPU-only for now.
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   */
  async groupByCount2(keysA, keysB) {
    const workloadSize = keysA.length;
    const backend = this._chooseBackend("groupByCount2", workloadSize, "u32");

    if (backend === "webgpu") {
      try {
        if (typeof this._gpu.groupByCount2 !== "function") {
          throw new Error("WebGPU backend does not implement groupByCount2");
        }
        const gpu = await this._gpu.groupByCount2(keysA, keysB);
        if (this._shouldValidate("groupByCount2", workloadSize)) {
          const cpu = await this._cpu.groupByCount2(keysA, keysB);
          let ok =
            cpu.uniqueKeysA.length === gpu.uniqueKeysA.length &&
            cpu.uniqueKeysB.length === gpu.uniqueKeysB.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.counts.length; i++) {
              if (cpu.uniqueKeysA[i] !== gpu.uniqueKeysA[i] || cpu.uniqueKeysB[i] !== gpu.uniqueKeysB[i] || cpu.counts[i] !== gpu.counts[i]) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByCount2", precision: "u32", workloadSize };
            this._lastKernelBackend.groupByCount2 = "cpu";
            this._lastKernelPrecision.groupByCount2 = "u32";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByCount2 = "webgpu";
        this._lastKernelPrecision.groupByCount2 = "u32";
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByCount2", "u32", err);
        this._lastKernelBackend.groupByCount2 = "cpu";
        this._lastKernelPrecision.groupByCount2 = "u32";
        return this._cpu.groupByCount2(keysA, keysB);
      }
    }

    this._lastKernelBackend.groupByCount2 = "cpu";
    this._lastKernelPrecision.groupByCount2 = "u32";
    return this._cpu.groupByCount2(keysA, keysB);
  }

  /**
   * Two-key group-by SUM(+COUNT). CPU-only for now.
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   */
  async groupBySum2(keysA, keysB, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keysA.length;
    const backend = this._chooseBackend("groupBySum2", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        if (typeof this._gpu.groupBySum2 !== "function") {
          throw new Error("WebGPU backend does not implement groupBySum2");
        }
        const gpu = await this._gpu.groupBySum2(keysA, keysB, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupBySum2", workloadSize)) {
          const cpu = await this._cpu.groupBySum2(keysA, keysB, values);
          let ok =
            cpu.uniqueKeysA.length === gpu.uniqueKeysA.length &&
            cpu.uniqueKeysB.length === gpu.uniqueKeysB.length &&
            cpu.sums.length === gpu.sums.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.counts.length; i++) {
              if (
                cpu.uniqueKeysA[i] !== gpu.uniqueKeysA[i] ||
                cpu.uniqueKeysB[i] !== gpu.uniqueKeysB[i] ||
                cpu.counts[i] !== gpu.counts[i] ||
                !this._withinTolerance(gpu.sums[i], cpu.sums[i])
              ) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupBySum2", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupBySum2 = "cpu";
            this._lastKernelPrecision.groupBySum2 = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupBySum2 = "webgpu";
        this._lastKernelPrecision.groupBySum2 = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupBySum2", gpuPrecision, err);
        this._lastKernelBackend.groupBySum2 = "cpu";
        this._lastKernelPrecision.groupBySum2 = "f64";
        return this._cpu.groupBySum2(keysA, keysB, values);
      }
    }

    this._lastKernelBackend.groupBySum2 = "cpu";
    this._lastKernelPrecision.groupBySum2 = "f64";
    return this._cpu.groupBySum2(keysA, keysB, values);
  }

  /**
   * Two-key group-by MIN(+COUNT). CPU-only for now.
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   */
  async groupByMin2(keysA, keysB, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keysA.length;
    const backend = this._chooseBackend("groupByMin2", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        if (typeof this._gpu.groupByMin2 !== "function") {
          throw new Error("WebGPU backend does not implement groupByMin2");
        }
        const gpu = await this._gpu.groupByMin2(keysA, keysB, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupByMin2", workloadSize)) {
          const cpu = await this._cpu.groupByMin2(keysA, keysB, values);
          let ok =
            cpu.uniqueKeysA.length === gpu.uniqueKeysA.length &&
            cpu.uniqueKeysB.length === gpu.uniqueKeysB.length &&
            cpu.mins.length === gpu.mins.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.counts.length; i++) {
              if (
                cpu.uniqueKeysA[i] !== gpu.uniqueKeysA[i] ||
                cpu.uniqueKeysB[i] !== gpu.uniqueKeysB[i] ||
                cpu.counts[i] !== gpu.counts[i] ||
                !Object.is(gpu.mins[i], cpu.mins[i])
              ) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByMin2", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupByMin2 = "cpu";
            this._lastKernelPrecision.groupByMin2 = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByMin2 = "webgpu";
        this._lastKernelPrecision.groupByMin2 = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByMin2", gpuPrecision, err);
        this._lastKernelBackend.groupByMin2 = "cpu";
        this._lastKernelPrecision.groupByMin2 = "f64";
        return this._cpu.groupByMin2(keysA, keysB, values);
      }
    }

    this._lastKernelBackend.groupByMin2 = "cpu";
    this._lastKernelPrecision.groupByMin2 = "f64";
    return this._cpu.groupByMin2(keysA, keysB, values);
  }

  /**
   * Two-key group-by MAX(+COUNT). CPU-only for now.
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   */
  async groupByMax2(keysA, keysB, values) {
    const gpuPrecision = this._gpuPrecisionForValues("f32", values);
    const workloadSize = keysA.length;
    const backend = this._chooseBackend("groupByMax2", workloadSize, gpuPrecision);

    if (backend === "webgpu") {
      try {
        if (typeof this._gpu.groupByMax2 !== "function") {
          throw new Error("WebGPU backend does not implement groupByMax2");
        }
        const gpu = await this._gpu.groupByMax2(keysA, keysB, values, {
          precision: gpuPrecision,
          allowFp32FallbackForF64: this._allowFp32FallbackForF64
        });
        if (this._shouldValidate("groupByMax2", workloadSize)) {
          const cpu = await this._cpu.groupByMax2(keysA, keysB, values);
          let ok =
            cpu.uniqueKeysA.length === gpu.uniqueKeysA.length &&
            cpu.uniqueKeysB.length === gpu.uniqueKeysB.length &&
            cpu.maxs.length === gpu.maxs.length &&
            cpu.counts.length === gpu.counts.length;
          if (ok) {
            for (let i = 0; i < cpu.counts.length; i++) {
              if (
                cpu.uniqueKeysA[i] !== gpu.uniqueKeysA[i] ||
                cpu.uniqueKeysB[i] !== gpu.uniqueKeysB[i] ||
                cpu.counts[i] !== gpu.counts[i] ||
                !Object.is(gpu.maxs[i], cpu.maxs[i])
              ) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "groupByMax2", precision: gpuPrecision, workloadSize };
            this._lastKernelBackend.groupByMax2 = "cpu";
            this._lastKernelPrecision.groupByMax2 = "f64";
            return cpu;
          }
        }
        this._lastKernelBackend.groupByMax2 = "webgpu";
        this._lastKernelPrecision.groupByMax2 = gpuPrecision;
        return gpu;
      } catch (err) {
        this._recordGpuError("groupByMax2", gpuPrecision, err);
        this._lastKernelBackend.groupByMax2 = "cpu";
        this._lastKernelPrecision.groupByMax2 = "f64";
        return this._cpu.groupByMax2(keysA, keysB, values);
      }
    }

    this._lastKernelBackend.groupByMax2 = "cpu";
    this._lastKernelPrecision.groupByMax2 = "f64";
    return this._cpu.groupByMax2(keysA, keysB, values);
  }

  /**
   * @param {Uint32Array | Int32Array} leftKeys
   * @param {Uint32Array | Int32Array} rightKeys
   * @param {{ joinType?: "inner" | "left" }} [opts]
   */
  async hashJoin(leftKeys, rightKeys, opts = {}) {
    if (leftKeys.length > 0 && rightKeys.length > 0) {
      const leftSigned = leftKeys instanceof Int32Array;
      const rightSigned = rightKeys instanceof Int32Array;
      if (leftSigned !== rightSigned) {
        throw new Error(
          `hashJoin key type mismatch: left=${leftSigned ? "i32" : "u32"} right=${rightSigned ? "i32" : "u32"} (pass matching Int32Array/Uint32Array types)`
        );
      }
    }

    const joinType = opts.joinType ?? "inner";
    if (joinType !== "inner" && joinType !== "left") {
      throw new Error(`hashJoin joinType must be "inner" | "left", got ${String(joinType)}`);
    }
    if (leftKeys.length === 0) {
      this._lastKernelBackend.hashJoin = "cpu";
      this._lastKernelPrecision.hashJoin = "u32";
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }
    if (rightKeys.length === 0) {
      this._lastKernelBackend.hashJoin = "cpu";
      this._lastKernelPrecision.hashJoin = "u32";
      if (joinType === "left") {
        const leftIndex = new Uint32Array(leftKeys.length);
        const rightIndex = new Uint32Array(leftKeys.length);
        for (let i = 0; i < leftKeys.length; i++) {
          leftIndex[i] = i;
          rightIndex[i] = 0xffff_ffff;
        }
        return { leftIndex, rightIndex };
      }
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }

    const workloadSize = leftKeys.length + rightKeys.length;
    const backend = this._chooseBackend("hashJoin", workloadSize, "u32");

    if (backend === "webgpu") {
      try {
        const gpu = await this._gpu.hashJoin(leftKeys, rightKeys, opts);
        if (this._shouldValidate("hashJoin", workloadSize)) {
          const cpu = await this._cpu.hashJoin(leftKeys, rightKeys, opts);
          let ok = cpu.leftIndex.length === gpu.leftIndex.length && cpu.rightIndex.length === gpu.rightIndex.length;
          if (ok) {
            for (let i = 0; i < cpu.leftIndex.length; i++) {
              if (cpu.leftIndex[i] !== gpu.leftIndex[i] || cpu.rightIndex[i] !== gpu.rightIndex[i]) {
                ok = false;
                break;
              }
            }
          }
          if (!ok) {
            this._validationState.mismatches += 1;
            this._validationState.lastMismatch = { kernel: "hashJoin", precision: "u32", workloadSize };
            this._lastKernelBackend.hashJoin = "cpu";
            this._lastKernelPrecision.hashJoin = "u32";
            return cpu;
          }
        }
        this._lastKernelBackend.hashJoin = "webgpu";
        this._lastKernelPrecision.hashJoin = "u32";
        return gpu;
      } catch (err) {
        this._recordGpuError("hashJoin", "u32", err);
        this._lastKernelBackend.hashJoin = "cpu";
        this._lastKernelPrecision.hashJoin = "u32";
        return this._cpu.hashJoin(leftKeys, rightKeys, opts);
      }
    }

    this._lastKernelBackend.hashJoin = "cpu";
    this._lastKernelPrecision.hashJoin = "u32";
    return this._cpu.hashJoin(leftKeys, rightKeys, opts);
  }
}

/**
 * Convenience wrapper for callers.
 * @param {KernelEngineOptions} opts
 */
export async function createKernelEngine(opts = {}) {
  return KernelEngine.create(opts);
}
