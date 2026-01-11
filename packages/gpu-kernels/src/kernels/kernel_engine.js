import { CpuBackend } from "./cpu_backend.js";
import { WebGpuBackend } from "./webgpu_backend.js";

export const DEFAULT_THRESHOLDS = {
  sum: 1 << 15,
  min: 1 << 15,
  max: 1 << 15,
  average: 1 << 15,
  count: 1 << 15,
  sumproduct: 1 << 15,
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
    this._validationState = { mismatches: 0, lastMismatch: null };

    /** @type {Record<string, "cpu" | "webgpu">} */
    this._lastKernelBackend = {
      sum: "cpu",
      min: "cpu",
      max: "cpu",
      average: "cpu",
      count: "cpu",
      sumproduct: "cpu",
      mmult: "cpu",
      sort: "cpu",
      histogram: "cpu"
    };

    /** @type {Record<string, "f32" | "f64">} */
    this._lastKernelPrecision = {
      sum: "f64",
      min: "f64",
      max: "f64",
      average: "f64",
      count: "f64",
      sumproduct: "f64",
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
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "mmult" | "sort" | "histogram"} kernel
   * @param {number} workloadSize
   * @param {"f32" | "f64"} gpuPrecision
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
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "mmult" | "sort" | "histogram"} kernel
   * @param {"f32" | "f64"} precision
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
   * @param {"sum" | "min" | "max" | "average" | "sumproduct" | "histogram"} kernel
   * @param {number} workloadSize
   */
  _shouldValidate(kernel, workloadSize) {
    if (!this._validation.enabled) return false;
    if (workloadSize > this._validation.maxElements) return false;
    return kernel === "sum" || kernel === "sumproduct" || kernel === "histogram" || kernel === "min" || kernel === "max" || kernel === "average";
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
      const gpu = await this._gpu.sum(values, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
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
      const gpu = await this._gpu.sumproduct(a, b, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
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
      this._lastKernelBackend.mmult = "webgpu";
      this._lastKernelPrecision.mmult = gpuPrecision;
      return this._gpu.mmult(a, b, aRows, aCols, bCols, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
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
      this._lastKernelBackend.sort = "webgpu";
      this._lastKernelPrecision.sort = gpuPrecision;
      return this._gpu.sort(values, { precision: gpuPrecision, allowFp32FallbackForF64: false });
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
      const gpu = await this._gpu.histogram(values, opts, { precision: gpuPrecision, allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
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
    }

    this._lastKernelBackend.histogram = "cpu";
    this._lastKernelPrecision.histogram = "f64";
    return this._cpu.histogram(values, opts);
  }
}

/**
 * Convenience wrapper for callers.
 * @param {KernelEngineOptions} opts
 */
export async function createKernelEngine(opts = {}) {
  return KernelEngine.create(opts);
}
