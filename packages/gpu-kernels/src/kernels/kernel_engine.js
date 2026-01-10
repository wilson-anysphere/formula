import { CpuBackend } from "./cpu_backend.js";
import { WebGpuBackend } from "./webgpu_backend.js";

export const DEFAULT_THRESHOLDS = {
  sum: 1 << 15,
  sumproduct: 1 << 15,
  // Rough heuristic based on multiply-add count (aRows * aCols * bCols).
  mmult: 1 << 20,
  sort: 1 << 15,
  histogram: 1 << 15
};

/**
 * @typedef {{
 *  gpu?: {
 *    enabled?: boolean,
 *    forceBackend?: "auto" | "cpu" | "gpu",
 *    allowFp32FallbackForF64?: boolean
 *  },
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
    this._gpuEnabled = opts.gpu?.enabled ?? true;
    this._forceBackend = opts.gpu?.forceBackend ?? "auto";
    this._allowFp32FallbackForF64 = opts.gpu?.allowFp32FallbackForF64 ?? true;
    this._thresholds = { ...DEFAULT_THRESHOLDS, ...(opts.thresholds ?? {}) };

    /** @type {Record<string, "cpu" | "webgpu">} */
    this._lastKernelBackend = {
      sum: "cpu",
      sumproduct: "cpu",
      mmult: "cpu",
      sort: "cpu",
      histogram: "cpu"
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
      gpu: {
        enabled: this._gpuEnabled,
        forceBackend: this._forceBackend,
        allowFp32FallbackForF64: this._allowFp32FallbackForF64,
        available: Boolean(this._gpu),
        ...(this._gpu ? this._gpu.diagnostics() : { kind: "webgpu", supportedKernels: {} })
      },
      cpu: this._cpu.diagnostics(),
      thresholds: { ...this._thresholds },
      lastKernelBackend: this.lastKernelBackend()
    };
  }

  async dispose() {
    if (this._gpu) this._gpu.dispose();
  }

  /**
   * @param {"sum" | "sumproduct" | "mmult" | "sort" | "histogram"} kernel
   * @param {number} workloadSize
   * @returns {"cpu" | "webgpu"}
   */
  _chooseBackend(kernel, workloadSize) {
    if (this._forceBackend === "cpu") return "cpu";
    if (this._forceBackend === "gpu") return this._gpu && this._gpuEnabled ? "webgpu" : "cpu";
    if (!this._gpu || !this._gpuEnabled) return "cpu";

    const threshold = this._thresholds[kernel] ?? Infinity;
    return workloadSize >= threshold ? "webgpu" : "cpu";
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async sum(values) {
    let backend = this._chooseBackend("sum", values.length);
    if (backend === "webgpu" && values instanceof Float64Array && !this._allowFp32FallbackForF64) {
      backend = "cpu";
    }

    if (backend === "webgpu") {
      this._lastKernelBackend.sum = "webgpu";
      return this._gpu.sum(values, { allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
    }
    this._lastKernelBackend.sum = "cpu";
    return this._cpu.sum(values);
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   */
  async sumproduct(a, b) {
    let backend = this._chooseBackend("sumproduct", a.length);
    if (
      backend === "webgpu" &&
      !this._allowFp32FallbackForF64 &&
      (a instanceof Float64Array || b instanceof Float64Array)
    ) {
      backend = "cpu";
    }

    if (backend === "webgpu") {
      this._lastKernelBackend.sumproduct = "webgpu";
      return this._gpu.sumproduct(a, b, { allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
    }
    this._lastKernelBackend.sumproduct = "cpu";
    return this._cpu.sumproduct(a, b);
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
    let backend = this._chooseBackend("mmult", mulAdds);
    if (
      backend === "webgpu" &&
      !this._allowFp32FallbackForF64 &&
      (a instanceof Float64Array || b instanceof Float64Array)
    ) {
      backend = "cpu";
    }

    if (backend === "webgpu") {
      this._lastKernelBackend.mmult = "webgpu";
      return this._gpu.mmult(a, b, aRows, aCols, bCols, { allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
    }
    this._lastKernelBackend.mmult = "cpu";
    return this._cpu.mmult(a, b, aRows, aCols, bCols);
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async sort(values) {
    // Sorting is order-sensitive; we avoid implicit f64->f32 downcast here.
    let backend = this._chooseBackend("sort", values.length);
    if (backend === "webgpu" && values instanceof Float64Array) {
      backend = "cpu";
    }

    if (backend === "webgpu") {
      this._lastKernelBackend.sort = "webgpu";
      return this._gpu.sort(values, { allowFp32FallbackForF64: false });
    }
    this._lastKernelBackend.sort = "cpu";
    return this._cpu.sort(values);
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ min: number, max: number, bins: number }} opts
   */
  async histogram(values, opts) {
    let backend = this._chooseBackend("histogram", values.length);
    if (backend === "webgpu" && values instanceof Float64Array && !this._allowFp32FallbackForF64) {
      backend = "cpu";
    }

    if (backend === "webgpu") {
      this._lastKernelBackend.histogram = "webgpu";
      return this._gpu.histogram(values, opts, { allowFp32FallbackForF64: this._allowFp32FallbackForF64 });
    }
    this._lastKernelBackend.histogram = "cpu";
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
