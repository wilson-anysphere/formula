import { loadTextResource } from "../util/load_text_resource.js";

/**
 * @typedef {"webgpu"} BackendKind
 */

const WORKGROUP_SIZE = 256;
const MMULT_TILE = 8;
const EMPTY_U32 = 0xffff_ffff;

function hasWebGpu() {
  return typeof navigator !== "undefined" && navigator && "gpu" in navigator && navigator.gpu;
}

function nextPowerOfTwo(n) {
  let p = 1;
  while (p < n) p *= 2;
  return p;
}

function dtypeOf(values) {
  if (values instanceof Float32Array) return "f32";
  if (values instanceof Float64Array) return "f64";
  throw new Error(`Unsupported array type: expected Float32Array or Float64Array`);
}

function toFloat32(values) {
  if (values instanceof Float32Array) return values;
  const out = new Float32Array(values.length);
  for (let i = 0; i < values.length; i++) out[i] = values[i];
  return out;
}

function toFloat64(values) {
  if (values instanceof Float64Array) return values;
  const out = new Float64Array(values.length);
  for (let i = 0; i < values.length; i++) out[i] = values[i];
  return out;
}

/**
 * @param {Uint32Array | Int32Array} keys
 */
function toUint32Keys(keys) {
  if (keys instanceof Uint32Array) return keys;
  if (keys instanceof Int32Array) return new Uint32Array(keys.buffer, keys.byteOffset, keys.length);
  throw new Error(`Unsupported key array type: expected Uint32Array or Int32Array`);
}

/**
 * @param {"f32" | "f64"} dtype
 */
function byteSizeOf(dtype) {
  return dtype === "f32" ? 4 : 8;
}

export class GpuVector {
  /**
   * @param {WebGpuBackend} backend
   * @param {GPUBuffer} buffer
   * @param {number} length
   * @param {"f32" | "f64"} dtype
   */
  constructor(backend, buffer, length, dtype) {
    this.backend = backend;
    this.buffer = buffer;
    this.length = length;
    this.dtype = dtype;
  }

  async toTypedArray() {
    return this.backend.readbackVector(this);
  }

  destroy() {
    this.buffer.destroy();
  }
}

export class WebGpuBackend {
  /** @type {BackendKind} */
  kind = "webgpu";

  /**
   * @param {GPUDevice} device
   * @param {GPUAdapter} adapter
   * @param {{
 *  supportsF64: boolean,
 *  pipelines: {
 *    reduceSum_f32: GPUComputePipeline,
 *    reduceMin_f32: GPUComputePipeline,
 *    reduceMax_f32: GPUComputePipeline,
 *    reduceSumproduct_f32: GPUComputePipeline,
 *    groupByClear: GPUComputePipeline,
 *    groupByCount: GPUComputePipeline,
 *    groupBySum_f32: GPUComputePipeline,
 *    groupByMin_f32: GPUComputePipeline,
 *    groupByMax_f32: GPUComputePipeline,
 *    hashJoinClear: GPUComputePipeline,
 *    hashJoinBuild: GPUComputePipeline,
 *    hashJoinCount: GPUComputePipeline,
 *    hashJoinFill: GPUComputePipeline,
 *    mmult_f32: GPUComputePipeline,
 *    histogram_f32: GPUComputePipeline,
 *    bitonicSort_f32: GPUComputePipeline,
 *    reduceSum_f64?: GPUComputePipeline,
 *    reduceMin_f64?: GPUComputePipeline,
 *    reduceMax_f64?: GPUComputePipeline,
 *    reduceSumproduct_f64?: GPUComputePipeline,
 *    mmult_f64?: GPUComputePipeline,
 *    histogram_f64?: GPUComputePipeline,
 *    bitonicSort_f64?: GPUComputePipeline
 *  }
 * }} resources
 */
  constructor(device, adapter, resources) {
    this.device = device;
    this.adapter = adapter;
    this.queue = device.queue;
    this.supportsF64 = resources.supportsF64;
    this.pipelines = resources.pipelines;
    this._disposed = false;
    /** @type {Set<"f32" | "f64">} */
    this._precisionUsed = new Set();
  }

  static async createIfSupported() {
    try {
      if (!hasWebGpu()) return null;

      /** @type {GPUAdapter | null} */
      const adapter = await navigator.gpu.requestAdapter({ powerPreference: "high-performance" });
      if (!adapter) return null;

      let supportsF64 = adapter.features?.has?.("shader-f64") ?? false;

      /** @type {GPUDevice} */
      let device;
      try {
        device = supportsF64 ? await adapter.requestDevice({ requiredFeatures: ["shader-f64"] }) : await adapter.requestDevice();
      } catch {
        // Fall back to an f32-only device if requesting shader-f64 fails for any
        // reason (browser quirks, partial implementations, etc).
        device = await adapter.requestDevice();
        supportsF64 = false;
      }
      supportsF64 = supportsF64 && (device.features?.has?.("shader-f64") ?? false);

      const [
        reduceSumSrc,
        reduceMinSrc,
        reduceMaxSrc,
        reduceSumproductSrc,
        mmultSrc,
        histogramSrc,
        bitonicSortSrc,
        groupByClearSrc,
        groupByCountSrc,
        groupBySumSrc,
        groupByMinSrc,
        groupByMaxSrc,
        hashJoinClearSrc,
        hashJoinBuildSrc,
        hashJoinCountSrc,
        hashJoinFillSrc
      ] = await Promise.all([
        loadTextResource(new URL("./wgsl/reduce_sum.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_min.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_max.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_sumproduct.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/mmult.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/histogram.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/bitonic_sort.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/groupby_clear.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/groupby_count.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/groupby_sum_f32.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/groupby_min_f32.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/groupby_max_f32.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/hash_join_clear.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/hash_join_build.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/hash_join_count.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/hash_join_fill.wgsl", import.meta.url))
      ]);

      const reduceSum_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceSumSrc }),
          entryPoint: "main"
        }
      });

      const reduceMin_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceMinSrc }),
          entryPoint: "main"
        }
      });

      const reduceMax_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceMaxSrc }),
          entryPoint: "main"
        }
      });

      const reduceSumproduct_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceSumproductSrc }),
          entryPoint: "main"
        }
      });

      const mmult_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: mmultSrc }),
          entryPoint: "main"
        }
      });

      const histogram_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: histogramSrc }),
          entryPoint: "main"
        }
      });

      const bitonicSort_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: bitonicSortSrc }),
          entryPoint: "main"
        }
      });

      const groupByClear = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: groupByClearSrc }),
          entryPoint: "main"
        }
      });

      const groupByCount = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: groupByCountSrc }),
          entryPoint: "main"
        }
      });

      const groupBySum_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: groupBySumSrc }),
          entryPoint: "main"
        }
      });

      const groupByMin_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: groupByMinSrc }),
          entryPoint: "main"
        }
      });

      const groupByMax_f32 = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: groupByMaxSrc }),
          entryPoint: "main"
        }
      });

      const hashJoinClear = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: hashJoinClearSrc }),
          entryPoint: "main"
        }
      });

      const hashJoinBuild = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: hashJoinBuildSrc }),
          entryPoint: "main"
        }
      });

      const hashJoinCount = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: hashJoinCountSrc }),
          entryPoint: "main"
        }
      });

      const hashJoinFill = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: hashJoinFillSrc }),
          entryPoint: "main"
        }
      });

      /** @type {WebGpuBackend["pipelines"]} */
      const pipelines = {
        reduceSum_f32,
        reduceMin_f32,
        reduceMax_f32,
        reduceSumproduct_f32,
        groupByClear,
        groupByCount,
        groupBySum_f32,
        groupByMin_f32,
        groupByMax_f32,
        hashJoinClear,
        hashJoinBuild,
        hashJoinCount,
        hashJoinFill,
        mmult_f32,
        histogram_f32,
        bitonicSort_f32
      };

      if (supportsF64) {
        // Create f64 variants. These only compile on devices that expose the
        // `shader-f64` feature. Keep the backend usable even if an individual
        // f64 shader/pipeline fails to compile by falling back to f32 for that
        // kernel.
        let reduceSumF64Src;
        let reduceMinF64Src;
        let reduceMaxF64Src;
        let reduceSumproductF64Src;
        let mmultF64Src;
        let histogramF64Src;
        let bitonicSortF64Src;
        try {
          [reduceSumF64Src, reduceMinF64Src, reduceMaxF64Src, reduceSumproductF64Src, mmultF64Src, histogramF64Src, bitonicSortF64Src] =
            await Promise.all([
              loadTextResource(new URL("./wgsl/reduce_sum_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/reduce_min_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/reduce_max_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/reduce_sumproduct_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/mmult_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/histogram_f64.wgsl", import.meta.url)),
              loadTextResource(new URL("./wgsl/bitonic_sort_f64.wgsl", import.meta.url))
            ]);
        } catch {
          // If fetching f64 WGSL fails, treat f64 as unavailable while keeping
          // f32 pipelines usable.
          supportsF64 = false;
        }

        if (supportsF64) {
          try {
            pipelines.reduceSum_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: reduceSumF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.reduceMin_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: reduceMinF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.reduceMax_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: reduceMaxF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.reduceSumproduct_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: reduceSumproductF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.mmult_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: mmultF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.histogram_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: histogramF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
          try {
            pipelines.bitonicSort_f64 = device.createComputePipeline({
              layout: "auto",
              compute: {
                module: device.createShaderModule({ code: bitonicSortF64Src }),
                entryPoint: "main"
              }
            });
          } catch {}
        }
      }

      return new WebGpuBackend(device, adapter, {
        supportsF64,
        pipelines
      });
    } catch {
      // WebGPU exists but the adapter/device/pipelines could not be created.
      // Treat this as "unsupported" and let callers fall back to CPU.
      return null;
    }
  }

  dispose() {
    this._disposed = true;
  }

  diagnostics() {
    return {
      kind: this.kind,
      adapterInfo: typeof this.adapter.requestAdapterInfo === "function" ? "available" : "unavailable",
      supportedKernels: {
        sum: true,
        min: true,
        max: true,
        sumproduct: true,
        average: true,
        count: true,
        groupByCount: true,
        groupBySum: true,
        groupByMin: true,
        groupByMax: true,
        hashJoin: true,
        mmult: true,
        sort: true,
        histogram: true
      },
      supportedKernelsF64: {
        sum: Boolean(this.pipelines.reduceSum_f64),
        min: Boolean(this.pipelines.reduceMin_f64),
        max: Boolean(this.pipelines.reduceMax_f64),
        sumproduct: Boolean(this.pipelines.reduceSumproduct_f64 && this.pipelines.reduceSum_f64),
        average: Boolean(this.pipelines.reduceSum_f64),
        count: Boolean(this.pipelines.reduceSum_f64),
        groupByCount: true,
        groupBySum: false,
        groupByMin: false,
        groupByMax: false,
        hashJoin: true,
        mmult: Boolean(this.pipelines.mmult_f64),
        sort: Boolean(this.pipelines.bitonicSort_f64),
        histogram: Boolean(this.pipelines.histogram_f64)
      },
      numericPrecision:
        this._precisionUsed.size === 0
          ? this.supportsF64
            ? "mixed"
            : "f32"
          : this._precisionUsed.size === 1
            ? Array.from(this._precisionUsed)[0]
            : "mixed",
      supportsF64: this.supportsF64,
      workgroupSize: WORKGROUP_SIZE
    };
  }

  /**
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "groupByCount" | "groupBySum" | "groupByMin" | "groupByMax" | "hashJoin" | "mmult" | "sort" | "histogram"} kernel
   * @param {"f32" | "f64" | "u32"} dtype
   */
  supportsKernelPrecision(kernel, dtype) {
    if (dtype === "f32") return true;
    if (dtype === "u32") {
      switch (kernel) {
        case "groupByCount":
          return Boolean(this.pipelines.groupByCount && this.pipelines.groupByClear);
        case "hashJoin":
          return Boolean(this.pipelines.hashJoinClear && this.pipelines.hashJoinBuild && this.pipelines.hashJoinCount && this.pipelines.hashJoinFill);
        default:
          return false;
      }
    }
    if (!this.supportsF64) return false;
    switch (kernel) {
      case "sum":
        return Boolean(this.pipelines.reduceSum_f64);
      case "min":
        return Boolean(this.pipelines.reduceMin_f64);
      case "max":
        return Boolean(this.pipelines.reduceMax_f64);
      case "sumproduct":
        // SUMPRODUCT uses its own first-pass kernel then finishes via the
        // regular sum reduction.
        return Boolean(this.pipelines.reduceSumproduct_f64 && this.pipelines.reduceSum_f64);
      case "groupBySum":
      case "groupByMin":
      case "groupByMax":
      case "groupByCount":
      case "hashJoin":
        return false;
      case "average":
      case "count":
        // `average` and `count` are derived from sum + scalar operations.
        return Boolean(this.pipelines.reduceSum_f64);
      case "mmult":
        return Boolean(this.pipelines.mmult_f64);
      case "sort":
        return Boolean(this.pipelines.bitonicSort_f64);
      case "histogram":
        return Boolean(this.pipelines.histogram_f64);
      default:
        return false;
    }
  }

  _ensureNotDisposed() {
    if (this._disposed) throw new Error("WebGpuBackend is disposed");
  }

  /**
   * WebGPU limits the number of workgroups per dimension. For very large inputs
   * we use a 2D dispatch and linearize indices in WGSL using `num_workgroups`.
   * @param {number} workgroups
   */
  _dispatch2D(workgroups) {
    const max = this.device.limits?.maxComputeWorkgroupsPerDimension ?? 65535;
    const x = Math.min(workgroups, max);
    const y = Math.ceil(workgroups / x);
    if (y > max) {
      throw new Error(`Dispatch requires ${workgroups} workgroups which exceeds device limits (${max}^2)`);
    }
    return { x, y, total: x * y };
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async sum(values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (dtype === "f64") {
      if (!this.supportsKernelPrecision("sum", "f64")) {
        if (!allowFp32FallbackForF64) {
          throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
        }
        const f32 = toFloat32(values);
        if (f32.length === 0) return 0;
        this._precisionUsed.add("f32");
        const vec = this.uploadVector(f32);
        try {
          return await this.sumVector(vec);
        } finally {
          vec.destroy();
        }
      }

      const f64 = toFloat64(values);
      if (f64.length === 0) return 0;
      this._precisionUsed.add("f64");
      const vec = this.uploadVector(f64);
      try {
        return await this.sumVector(vec);
      } finally {
        vec.destroy();
      }
    }

    const f32 = toFloat32(values);
    if (f32.length === 0) return 0;
    this._precisionUsed.add("f32");
    const vec = this.uploadVector(f32);
    try {
      return await this.sumVector(vec);
    } finally {
      vec.destroy();
    }
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async min(values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (values.length === 0) return Number.POSITIVE_INFINITY;

    if (dtype === "f64") {
      if (!this.supportsKernelPrecision("min", "f64")) {
        if (!allowFp32FallbackForF64) {
          throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
        }
        const f32 = toFloat32(values);
        this._precisionUsed.add("f32");
        const vec = this.uploadVector(f32);
        try {
          return await this.minVector(vec);
        } finally {
          vec.destroy();
        }
      }

      const f64 = toFloat64(values);
      this._precisionUsed.add("f64");
      const vec = this.uploadVector(f64);
      try {
        return await this.minVector(vec);
      } finally {
        vec.destroy();
      }
    }

    const f32 = toFloat32(values);
    this._precisionUsed.add("f32");
    const vec = this.uploadVector(f32);
    try {
      return await this.minVector(vec);
    } finally {
      vec.destroy();
    }
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async max(values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (values.length === 0) return Number.NEGATIVE_INFINITY;

    if (dtype === "f64") {
      if (!this.supportsKernelPrecision("max", "f64")) {
        if (!allowFp32FallbackForF64) {
          throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
        }
        const f32 = toFloat32(values);
        this._precisionUsed.add("f32");
        const vec = this.uploadVector(f32);
        try {
          return await this.maxVector(vec);
        } finally {
          vec.destroy();
        }
      }

      const f64 = toFloat64(values);
      this._precisionUsed.add("f64");
      const vec = this.uploadVector(f64);
      try {
        return await this.maxVector(vec);
      } finally {
        vec.destroy();
      }
    }

    const f32 = toFloat32(values);
    this._precisionUsed.add("f32");
    const vec = this.uploadVector(f32);
    try {
      return await this.maxVector(vec);
    } finally {
      vec.destroy();
    }
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async average(values, opts) {
    if (values.length === 0) return Number.NaN;
    const total = await this.sum(values, opts);
    return total / values.length;
  }

  /**
   * Numeric-only count. For typed arrays this is equivalent to the array length.
   * @param {Float32Array | Float64Array} values
   */
  async count(values) {
    return values.length;
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async sumproduct(a, b, opts) {
    this._ensureNotDisposed();
    if (a.length !== b.length) throw new Error(`SUMPRODUCT length mismatch: ${a.length} vs ${b.length}`);

    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const aDtype = dtypeOf(a);
    const bDtype = dtypeOf(b);
    const inferred = aDtype === "f64" || bDtype === "f64" ? "f64" : "f32";
    const dtype = requested === "auto" ? inferred : requested;

    if (dtype === "f64") {
      if (!this.supportsKernelPrecision("sumproduct", "f64")) {
        if (!allowFp32FallbackForF64) {
          throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
        }
        const a32 = toFloat32(a);
        const b32 = toFloat32(b);
        if (a32.length === 0) return 0;
        this._precisionUsed.add("f32");
        const ga = this.uploadVector(a32);
        const gb = this.uploadVector(b32);
        try {
          return await this.sumproductVectors(ga, gb);
        } finally {
          ga.destroy();
          gb.destroy();
        }
      }

      const a64 = toFloat64(a);
      const b64 = toFloat64(b);
      if (a64.length === 0) return 0;
      this._precisionUsed.add("f64");
      const ga = this.uploadVector(a64);
      const gb = this.uploadVector(b64);
      try {
        return await this.sumproductVectors(ga, gb);
      } finally {
        ga.destroy();
        gb.destroy();
      }
    }

    const a32 = toFloat32(a);
    const b32 = toFloat32(b);
    if (a32.length === 0) return 0;
    this._precisionUsed.add("f32");
    const ga = this.uploadVector(a32);
    const gb = this.uploadVector(b32);
    try {
      return await this.sumproductVectors(ga, gb);
    } finally {
      ga.destroy();
      gb.destroy();
    }
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   * @param {number} aRows
   * @param {number} aCols
   * @param {number} bCols
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async mmult(a, b, aRows, aCols, bCols, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";

    if (a.length !== aRows * aCols) {
      throw new Error(`MMULT A shape mismatch: a.length=${a.length} vs ${aRows}x${aCols}`);
    }
    if (b.length !== aCols * bCols) {
      throw new Error(`MMULT B shape mismatch: b.length=${b.length} vs ${aCols}x${bCols}`);
    }
    if (aRows === 0 || bCols === 0) {
      return new Float64Array();
    }
    if (aCols === 0) {
      // 0-column matrix multiplication produces a zero matrix output.
      return new Float64Array(aRows * bCols);
    }

    const aDtype = dtypeOf(a);
    const bDtype = dtypeOf(b);
    const inferred = aDtype === "f64" || bDtype === "f64" ? "f64" : "f32";
    const dtype = requested === "auto" ? inferred : requested;

    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("mmult", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
    }

    const pipeline = canRunF64 ? this.pipelines.mmult_f64 : this.pipelines.mmult_f32;
    const aTyped = canRunF64 ? toFloat64(a) : toFloat32(a);
    const bTyped = canRunF64 ? toFloat64(b) : toFloat32(b);

    const aBuf = this._createStorageBufferFromArray(aTyped, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const bBuf = this._createStorageBufferFromArray(bTyped, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const outBytes = aRows * bCols * byteSizeOf(canRunF64 ? "f64" : "f32");
    const outBuf = this.device.createBuffer({
      size: outBytes,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
    });

    const params = new Uint32Array([aRows, aCols, bCols, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: pipeline.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: aBuf } },
        { binding: 1, resource: { buffer: bBuf } },
        { binding: 2, resource: { buffer: outBuf } },
        { binding: 3, resource: { buffer: paramBuf } }
      ]
    });

    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(pipeline);
    pass.setBindGroup(0, bindGroup);
    const wgX = Math.ceil(bCols / MMULT_TILE);
    const wgY = Math.ceil(aRows / MMULT_TILE);
    const max = this.device.limits?.maxComputeWorkgroupsPerDimension ?? 65535;
    if (wgX > max || wgY > max) {
      throw new Error(`MMULT dispatch exceeds device limits: ${wgX}x${wgY} workgroups (max ${max})`);
    }
    pass.dispatchWorkgroups(wgX, wgY, 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    const out = await this._readbackTypedArray(outBuf, canRunF64 ? Float64Array : Float32Array);

    aBuf.destroy();
    bBuf.destroy();
    outBuf.destroy();
    paramBuf.destroy();

    this._precisionUsed.add(canRunF64 ? "f64" : "f32");
    return canRunF64 ? out : Float64Array.from(out);
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   */
  async sort(values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("sort", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
    }

    const typed = canRunF64 ? toFloat64(values) : toFloat32(values);
    const n = typed.length;
    if (n === 0) return new Float64Array();
    const padded = nextPowerOfTwo(n);
    const data = canRunF64 ? new Float64Array(padded) : new Float32Array(padded);
    data.set(typed);
    if (padded !== n) data.fill(Number.NaN, n);

    const dataBuf = this._createStorageBufferFromArray(data, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC);

    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });

    const bindGroup = this.device.createBindGroup({
      layout: (canRunF64 ? this.pipelines.bitonicSort_f64 : this.pipelines.bitonicSort_f32).getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: dataBuf } },
        { binding: 1, resource: { buffer: paramBuf } }
      ]
    });

    const dispatch = this._dispatch2D(Math.ceil(padded / WORKGROUP_SIZE));
    for (let k = 2; k <= padded; k <<= 1) {
      for (let j = k >> 1; j > 0; j >>= 1) {
        const params = new Uint32Array([padded, j, k, 0]);
        this.queue.writeBuffer(paramBuf, 0, params);

        const encoder = this.device.createCommandEncoder();
        const pass = encoder.beginComputePass();
        pass.setPipeline(canRunF64 ? this.pipelines.bitonicSort_f64 : this.pipelines.bitonicSort_f32);
        pass.setBindGroup(0, bindGroup);
        pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
        pass.end();
        this.queue.submit([encoder.finish()]);
      }
    }

    const outPadded = await this._readbackTypedArray(dataBuf, canRunF64 ? Float64Array : Float32Array);
    dataBuf.destroy();
    paramBuf.destroy();

    this._precisionUsed.add(canRunF64 ? "f64" : "f32");
    return canRunF64 ? outPadded.subarray(0, n) : Float64Array.from(outPadded.subarray(0, n));
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ min: number, max: number, bins: number }} opts
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} backendOpts
   */
  async histogram(values, opts, backendOpts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = backendOpts.allowFp32FallbackForF64 ?? true;
    const requested = backendOpts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;
    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("histogram", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 kernels; disallowing f64->f32 fallback");
    }

    const typed = canRunF64 ? toFloat64(values) : toFloat32(values);
    const { min, max, bins } = opts;
    if (!(bins > 0)) throw new Error("histogram bins must be > 0");
    if (!(max > min)) throw new Error("histogram max must be > min");

    if (typed.length === 0) {
      return new Uint32Array(bins);
    }

    const inputBuf = this._createStorageBufferFromArray(typed, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);

    const binBuf = this.device.createBuffer({
      size: bins * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(binBuf, 0, new Uint32Array(bins));

    const invBinWidth = bins / (max - min);
    const paramBufSize = canRunF64 ? 48 : 32;
    const params = new ArrayBuffer(paramBufSize);
    const view = new DataView(params);
    if (canRunF64) {
      view.setFloat64(0, min, true);
      view.setFloat64(8, max, true);
      view.setFloat64(16, invBinWidth, true);
      view.setUint32(24, typed.length, true);
      view.setUint32(28, bins, true);
    } else {
      view.setFloat32(0, min, true);
      view.setFloat32(4, max, true);
      view.setFloat32(8, invBinWidth, true);
      view.setUint32(12, typed.length, true);
      view.setUint32(16, bins, true);
    }

    const paramBuf = this.device.createBuffer({
      size: paramBufSize,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: (canRunF64 ? this.pipelines.histogram_f64 : this.pipelines.histogram_f32).getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inputBuf } },
        { binding: 1, resource: { buffer: binBuf } },
        { binding: 2, resource: { buffer: paramBuf } }
      ]
    });

    const dispatch = this._dispatch2D(Math.ceil(typed.length / WORKGROUP_SIZE));
    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(canRunF64 ? this.pipelines.histogram_f64 : this.pipelines.histogram_f32);
    pass.setBindGroup(0, bindGroup);
    pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    const counts = await this._readbackTypedArray(binBuf, Uint32Array);
    inputBuf.destroy();
    binBuf.destroy();
    paramBuf.destroy();
    this._precisionUsed.add(canRunF64 ? "f64" : "f32");
    return counts;
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, counts: Uint32Array }>}
   */
  async groupByCount(keys) {
    this._ensureNotDisposed();

    const signedKeys = keys instanceof Int32Array;
    const keysU32 = toUint32Keys(keys);
    const n = keysU32.length;
    if (n === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), counts: new Uint32Array() };
    }

    // Load factor ~0.5 (2x) keeps probe lengths manageable.
    const tableSize = nextPowerOfTwo(n * 2);
    const tableLen = tableSize + 1; // +1 for the special key (0xFFFF_FFFF).

    const inKeysBuf = this._createStorageBufferFromArray(keysU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const tableKeysBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableCountsBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    // Unused, but the clear kernel expects an aggregate buffer.
    const tableAggBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    const clearParams = new Uint32Array([tableLen, 0, 0, 0]);
    const clearParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(clearParamBuf, 0, clearParams);

    const clearBindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByClear.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: tableKeysBuf } },
        { binding: 1, resource: { buffer: tableCountsBuf } },
        { binding: 2, resource: { buffer: tableAggBuf } },
        { binding: 3, resource: { buffer: clearParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(tableLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByClear);
      pass.setBindGroup(0, clearBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const params = new Uint32Array([n, tableSize, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByCount.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inKeysBuf } },
        { binding: 1, resource: { buffer: tableKeysBuf } },
        { binding: 2, resource: { buffer: tableCountsBuf } },
        { binding: 3, resource: { buffer: paramBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(n / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByCount);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const [tableKeys, tableCounts] = await Promise.all([
      this._readbackTypedArray(tableKeysBuf, Uint32Array),
      this._readbackTypedArray(tableCountsBuf, Uint32Array)
    ]);

    inKeysBuf.destroy();
    tableKeysBuf.destroy();
    tableCountsBuf.destroy();
    tableAggBuf.destroy();
    clearParamBuf.destroy();
    paramBuf.destroy();

    /** @type {number[]} */
    const compactKeys = [];
    /** @type {number[]} */
    const compactCounts = [];
    for (let i = 0; i < tableLen; i++) {
      const c = tableCounts[i];
      if (c === 0) continue;
      compactKeys.push(tableKeys[i]);
      compactCounts.push(c);
    }

    if (compactKeys.length === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), counts: new Uint32Array() };
    }

    // Sort by key, then materialize typed arrays.
    const packed = new BigUint64Array(compactKeys.length);
    for (let i = 0; i < compactKeys.length; i++) {
      const sortKey = (compactKeys[i] ^ (signedKeys ? 0x8000_0000 : 0)) >>> 0;
      packed[i] = (BigInt(sortKey) << 32n) | BigInt(i);
    }
    packed.sort();

    const uniqueKeys = signedKeys ? new Int32Array(compactKeys.length) : new Uint32Array(compactKeys.length);
    const counts = new Uint32Array(compactKeys.length);
    for (let out = 0; out < packed.length; out++) {
      const idx = Number(packed[out] & 0xffff_ffffn);
      const keyU32 = compactKeys[idx] >>> 0;
      uniqueKeys[out] = signedKeys ? (keyU32 | 0) : keyU32;
      counts[out] = compactCounts[idx];
    }

    return { uniqueKeys, counts };
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, sums: Float64Array, counts: Uint32Array }>}
   */
  async groupBySum(keys, values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (keys.length !== values.length) {
      throw new Error(`groupBySum length mismatch: keys=${keys.length} values=${values.length}`);
    }

    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("groupBySum", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 group-by kernels; disallowing f64->f32 fallback");
    }

    const signedKeys = keys instanceof Int32Array;
    const keysU32 = toUint32Keys(keys);
    const vals = canRunF64 ? toFloat64(values) : toFloat32(values);
    const n = keysU32.length;
    if (n === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), sums: new Float64Array(), counts: new Uint32Array() };
    }

    const tableSize = nextPowerOfTwo(n * 2);
    const tableLen = tableSize + 1;

    const inKeysBuf = this._createStorageBufferFromArray(keysU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const inValsBuf = this._createStorageBufferFromArray(vals, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const tableKeysBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableCountsBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableAggBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    const clearParams = new Uint32Array([tableLen, 0 /* aggInitBits */, 0, 0]);
    const clearParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(clearParamBuf, 0, clearParams);

    const clearBindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByClear.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: tableKeysBuf } },
        { binding: 1, resource: { buffer: tableCountsBuf } },
        { binding: 2, resource: { buffer: tableAggBuf } },
        { binding: 3, resource: { buffer: clearParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(tableLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByClear);
      pass.setBindGroup(0, clearBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const params = new Uint32Array([n, tableSize, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupBySum_f32.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inKeysBuf } },
        { binding: 1, resource: { buffer: inValsBuf } },
        { binding: 2, resource: { buffer: tableKeysBuf } },
        { binding: 3, resource: { buffer: tableCountsBuf } },
        { binding: 4, resource: { buffer: tableAggBuf } },
        { binding: 5, resource: { buffer: paramBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(n / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupBySum_f32);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const [tableKeys, tableCounts, tableAggBits] = await Promise.all([
      this._readbackTypedArray(tableKeysBuf, Uint32Array),
      this._readbackTypedArray(tableCountsBuf, Uint32Array),
      this._readbackTypedArray(tableAggBuf, Uint32Array)
    ]);

    inKeysBuf.destroy();
    inValsBuf.destroy();
    tableKeysBuf.destroy();
    tableCountsBuf.destroy();
    tableAggBuf.destroy();
    clearParamBuf.destroy();
    paramBuf.destroy();

    const tableAgg = new Float32Array(tableAggBits.buffer, tableAggBits.byteOffset, tableAggBits.length);

    /** @type {number[]} */
    const compactKeys = [];
    /** @type {number[]} */
    const compactCounts = [];
    /** @type {number[]} */
    const compactSums = [];
    for (let i = 0; i < tableLen; i++) {
      const c = tableCounts[i];
      if (c === 0) continue;
      compactKeys.push(tableKeys[i]);
      compactCounts.push(c);
      compactSums.push(tableAgg[i]);
    }

    if (compactKeys.length === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), sums: new Float64Array(), counts: new Uint32Array() };
    }

    const packed = new BigUint64Array(compactKeys.length);
    for (let i = 0; i < compactKeys.length; i++) {
      const sortKey = (compactKeys[i] ^ (signedKeys ? 0x8000_0000 : 0)) >>> 0;
      packed[i] = (BigInt(sortKey) << 32n) | BigInt(i);
    }
    packed.sort();

    const uniqueKeys = signedKeys ? new Int32Array(compactKeys.length) : new Uint32Array(compactKeys.length);
    const counts = new Uint32Array(compactKeys.length);
    const sums = new Float64Array(compactKeys.length);
    for (let out = 0; out < packed.length; out++) {
      const idx = Number(packed[out] & 0xffff_ffffn);
      const keyU32 = compactKeys[idx] >>> 0;
      uniqueKeys[out] = signedKeys ? (keyU32 | 0) : keyU32;
      counts[out] = compactCounts[idx];
      sums[out] = compactSums[idx];
    }

    this._precisionUsed.add("f32");
    return { uniqueKeys, sums, counts };
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, mins: Float64Array, counts: Uint32Array }>}
   */
  async groupByMin(keys, values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (keys.length !== values.length) {
      throw new Error(`groupByMin length mismatch: keys=${keys.length} values=${values.length}`);
    }

    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("groupByMin", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 group-by kernels; disallowing f64->f32 fallback");
    }

    const signedKeys = keys instanceof Int32Array;
    const keysU32 = toUint32Keys(keys);
    const vals = canRunF64 ? toFloat64(values) : toFloat32(values);
    const n = keysU32.length;
    if (n === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), mins: new Float64Array(), counts: new Uint32Array() };
    }

    const tableSize = nextPowerOfTwo(n * 2);
    const tableLen = tableSize + 1;

    const inKeysBuf = this._createStorageBufferFromArray(keysU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const inValsBuf = this._createStorageBufferFromArray(vals, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const tableKeysBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableCountsBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableAggBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    const posInfBits = new Uint32Array(new Float32Array([Number.POSITIVE_INFINITY]).buffer)[0];
    const clearParams = new Uint32Array([tableLen, posInfBits, 0, 0]);
    const clearParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(clearParamBuf, 0, clearParams);

    const clearBindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByClear.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: tableKeysBuf } },
        { binding: 1, resource: { buffer: tableCountsBuf } },
        { binding: 2, resource: { buffer: tableAggBuf } },
        { binding: 3, resource: { buffer: clearParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(tableLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByClear);
      pass.setBindGroup(0, clearBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const params = new Uint32Array([n, tableSize, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByMin_f32.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inKeysBuf } },
        { binding: 1, resource: { buffer: inValsBuf } },
        { binding: 2, resource: { buffer: tableKeysBuf } },
        { binding: 3, resource: { buffer: tableCountsBuf } },
        { binding: 4, resource: { buffer: tableAggBuf } },
        { binding: 5, resource: { buffer: paramBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(n / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByMin_f32);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const [tableKeys, tableCounts, tableAggBits] = await Promise.all([
      this._readbackTypedArray(tableKeysBuf, Uint32Array),
      this._readbackTypedArray(tableCountsBuf, Uint32Array),
      this._readbackTypedArray(tableAggBuf, Uint32Array)
    ]);

    inKeysBuf.destroy();
    inValsBuf.destroy();
    tableKeysBuf.destroy();
    tableCountsBuf.destroy();
    tableAggBuf.destroy();
    clearParamBuf.destroy();
    paramBuf.destroy();

    const tableAgg = new Float32Array(tableAggBits.buffer, tableAggBits.byteOffset, tableAggBits.length);

    /** @type {number[]} */
    const compactKeys = [];
    /** @type {number[]} */
    const compactCounts = [];
    /** @type {number[]} */
    const compactMins = [];
    for (let i = 0; i < tableLen; i++) {
      const c = tableCounts[i];
      if (c === 0) continue;
      compactKeys.push(tableKeys[i]);
      compactCounts.push(c);
      compactMins.push(tableAgg[i]);
    }

    if (compactKeys.length === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), mins: new Float64Array(), counts: new Uint32Array() };
    }

    const packed = new BigUint64Array(compactKeys.length);
    for (let i = 0; i < compactKeys.length; i++) {
      const sortKey = (compactKeys[i] ^ (signedKeys ? 0x8000_0000 : 0)) >>> 0;
      packed[i] = (BigInt(sortKey) << 32n) | BigInt(i);
    }
    packed.sort();

    const uniqueKeys = signedKeys ? new Int32Array(compactKeys.length) : new Uint32Array(compactKeys.length);
    const counts = new Uint32Array(compactKeys.length);
    const mins = new Float64Array(compactKeys.length);
    for (let out = 0; out < packed.length; out++) {
      const idx = Number(packed[out] & 0xffff_ffffn);
      const keyU32 = compactKeys[idx] >>> 0;
      uniqueKeys[out] = signedKeys ? (keyU32 | 0) : keyU32;
      counts[out] = compactCounts[idx];
      mins[out] = compactMins[idx];
    }

    this._precisionUsed.add("f32");
    return { uniqueKeys, mins, counts };
  }

  /**
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64?: boolean, precision?: "auto" | "f32" | "f64" }} opts
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, maxs: Float64Array, counts: Uint32Array }>}
   */
  async groupByMax(keys, values, opts) {
    this._ensureNotDisposed();
    const allowFp32FallbackForF64 = opts.allowFp32FallbackForF64 ?? true;
    const requested = opts.precision ?? "auto";
    const dtype = requested === "auto" ? dtypeOf(values) : requested;

    if (keys.length !== values.length) {
      throw new Error(`groupByMax length mismatch: keys=${keys.length} values=${values.length}`);
    }

    const canRunF64 = dtype === "f64" && this.supportsKernelPrecision("groupByMax", "f64");
    if (dtype === "f64" && !canRunF64 && !allowFp32FallbackForF64) {
      throw new Error("WebGPU backend does not support f64 group-by kernels; disallowing f64->f32 fallback");
    }

    const signedKeys = keys instanceof Int32Array;
    const keysU32 = toUint32Keys(keys);
    const vals = canRunF64 ? toFloat64(values) : toFloat32(values);
    const n = keysU32.length;
    if (n === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), maxs: new Float64Array(), counts: new Uint32Array() };
    }

    const tableSize = nextPowerOfTwo(n * 2);
    const tableLen = tableSize + 1;

    const inKeysBuf = this._createStorageBufferFromArray(keysU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const inValsBuf = this._createStorageBufferFromArray(vals, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const tableKeysBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableCountsBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableAggBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    const negInfBits = new Uint32Array(new Float32Array([Number.NEGATIVE_INFINITY]).buffer)[0];
    const clearParams = new Uint32Array([tableLen, negInfBits, 0, 0]);
    const clearParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(clearParamBuf, 0, clearParams);

    const clearBindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByClear.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: tableKeysBuf } },
        { binding: 1, resource: { buffer: tableCountsBuf } },
        { binding: 2, resource: { buffer: tableAggBuf } },
        { binding: 3, resource: { buffer: clearParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(tableLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByClear);
      pass.setBindGroup(0, clearBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const params = new Uint32Array([n, tableSize, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.groupByMax_f32.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inKeysBuf } },
        { binding: 1, resource: { buffer: inValsBuf } },
        { binding: 2, resource: { buffer: tableKeysBuf } },
        { binding: 3, resource: { buffer: tableCountsBuf } },
        { binding: 4, resource: { buffer: tableAggBuf } },
        { binding: 5, resource: { buffer: paramBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(n / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.groupByMax_f32);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const [tableKeys, tableCounts, tableAggBits] = await Promise.all([
      this._readbackTypedArray(tableKeysBuf, Uint32Array),
      this._readbackTypedArray(tableCountsBuf, Uint32Array),
      this._readbackTypedArray(tableAggBuf, Uint32Array)
    ]);

    inKeysBuf.destroy();
    inValsBuf.destroy();
    tableKeysBuf.destroy();
    tableCountsBuf.destroy();
    tableAggBuf.destroy();
    clearParamBuf.destroy();
    paramBuf.destroy();

    const tableAgg = new Float32Array(tableAggBits.buffer, tableAggBits.byteOffset, tableAggBits.length);

    /** @type {number[]} */
    const compactKeys = [];
    /** @type {number[]} */
    const compactCounts = [];
    /** @type {number[]} */
    const compactMaxs = [];
    for (let i = 0; i < tableLen; i++) {
      const c = tableCounts[i];
      if (c === 0) continue;
      compactKeys.push(tableKeys[i]);
      compactCounts.push(c);
      compactMaxs.push(tableAgg[i]);
    }

    if (compactKeys.length === 0) {
      return { uniqueKeys: signedKeys ? new Int32Array() : new Uint32Array(), maxs: new Float64Array(), counts: new Uint32Array() };
    }

    const packed = new BigUint64Array(compactKeys.length);
    for (let i = 0; i < compactKeys.length; i++) {
      const sortKey = (compactKeys[i] ^ (signedKeys ? 0x8000_0000 : 0)) >>> 0;
      packed[i] = (BigInt(sortKey) << 32n) | BigInt(i);
    }
    packed.sort();

    const uniqueKeys = signedKeys ? new Int32Array(compactKeys.length) : new Uint32Array(compactKeys.length);
    const counts = new Uint32Array(compactKeys.length);
    const maxs = new Float64Array(compactKeys.length);
    for (let out = 0; out < packed.length; out++) {
      const idx = Number(packed[out] & 0xffff_ffffn);
      const keyU32 = compactKeys[idx] >>> 0;
      uniqueKeys[out] = signedKeys ? (keyU32 | 0) : keyU32;
      counts[out] = compactCounts[idx];
      maxs[out] = compactMaxs[idx];
    }

    this._precisionUsed.add("f32");
    return { uniqueKeys, maxs, counts };
  }

  /**
   * @param {Uint32Array | Int32Array} leftKeys
   * @param {Uint32Array | Int32Array} rightKeys
   * @param {{ joinType?: "inner" | "left" }} [opts]
   * @returns {Promise<{ leftIndex: Uint32Array, rightIndex: Uint32Array }>}
   */
  async hashJoin(leftKeys, rightKeys, opts = {}) {
    this._ensureNotDisposed();

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
    const joinTypeU32 = joinType === "left" ? 1 : 0;

    const leftU32 = toUint32Keys(leftKeys);
    const rightU32 = toUint32Keys(rightKeys);
    const leftLen = leftU32.length;
    const rightLen = rightU32.length;
    if (leftLen === 0) {
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }
    if (rightLen === 0) {
      if (joinType === "left") {
        const leftIndex = new Uint32Array(leftLen);
        const rightIndex = new Uint32Array(leftLen);
        for (let i = 0; i < leftLen; i++) {
          leftIndex[i] = i;
          rightIndex[i] = EMPTY_U32;
        }
        return { leftIndex, rightIndex };
      }
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }

    const tableSize = nextPowerOfTwo(rightLen * 2);
    const tableLen = tableSize + 1;

    const leftKeysBuf = this._createStorageBufferFromArray(leftU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const rightKeysBuf = this._createStorageBufferFromArray(rightU32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);

    const tableKeysBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const tableHeadsBuf = this.device.createBuffer({
      size: tableLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    const nextBuf = this.device.createBuffer({
      size: rightLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    // Clear hash table.
    const clearParams = new Uint32Array([tableLen, 0, 0, 0]);
    const clearParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(clearParamBuf, 0, clearParams);

    const clearBindGroup = this.device.createBindGroup({
      layout: this.pipelines.hashJoinClear.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: tableKeysBuf } },
        { binding: 1, resource: { buffer: tableHeadsBuf } },
        { binding: 2, resource: { buffer: clearParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(tableLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.hashJoinClear);
      pass.setBindGroup(0, clearBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    // Build hash table from right keys.
    const buildParams = new Uint32Array([rightLen, tableSize, 0, 0]);
    const buildParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(buildParamBuf, 0, buildParams);

    const buildBindGroup = this.device.createBindGroup({
      layout: this.pipelines.hashJoinBuild.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: rightKeysBuf } },
        { binding: 1, resource: { buffer: tableKeysBuf } },
        { binding: 2, resource: { buffer: tableHeadsBuf } },
        { binding: 3, resource: { buffer: nextBuf } },
        { binding: 4, resource: { buffer: buildParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(rightLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.hashJoinBuild);
      pass.setBindGroup(0, buildBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    // Count matches for each left key.
    const countsBuf = this.device.createBuffer({
      size: leftLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });

    const countParams = new Uint32Array([leftLen, tableSize, joinTypeU32, 0]);
    const countParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(countParamBuf, 0, countParams);

    const countBindGroup = this.device.createBindGroup({
      layout: this.pipelines.hashJoinCount.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: leftKeysBuf } },
        { binding: 1, resource: { buffer: tableKeysBuf } },
        { binding: 2, resource: { buffer: tableHeadsBuf } },
        { binding: 3, resource: { buffer: nextBuf } },
        { binding: 4, resource: { buffer: countsBuf } },
        { binding: 5, resource: { buffer: countParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(leftLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.hashJoinCount);
      pass.setBindGroup(0, countBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const counts = await this._readbackTypedArray(countsBuf, Uint32Array);
    const offsets = new Uint32Array(leftLen);
    let total = 0;
    for (let i = 0; i < leftLen; i++) {
      offsets[i] = total;
      total += counts[i];
      // Guard against overflow; outputs are Uint32-indexed.
      if (total > 0xffff_ffff) {
        throw new Error(`hashJoin output too large: ${total} pairs (exceeds Uint32Array limits)`);
      }
    }

    if (total === 0) {
      leftKeysBuf.destroy();
      rightKeysBuf.destroy();
      tableKeysBuf.destroy();
      tableHeadsBuf.destroy();
      nextBuf.destroy();
      clearParamBuf.destroy();
      buildParamBuf.destroy();
      countParamBuf.destroy();
      countsBuf.destroy();
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }

    const offsetsBuf = this._createStorageBufferFromArray(offsets, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const outLeftBuf = this.device.createBuffer({
      size: total * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
    });
    const outRightBuf = this.device.createBuffer({
      size: total * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
    });

    const fillParams = new Uint32Array([leftLen, tableSize, joinTypeU32, 0]);
    const fillParamBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(fillParamBuf, 0, fillParams);

    const fillBindGroup = this.device.createBindGroup({
      layout: this.pipelines.hashJoinFill.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: leftKeysBuf } },
        { binding: 1, resource: { buffer: tableKeysBuf } },
        { binding: 2, resource: { buffer: tableHeadsBuf } },
        { binding: 3, resource: { buffer: nextBuf } },
        { binding: 4, resource: { buffer: countsBuf } },
        { binding: 5, resource: { buffer: offsetsBuf } },
        { binding: 6, resource: { buffer: outLeftBuf } },
        { binding: 7, resource: { buffer: outRightBuf } },
        { binding: 8, resource: { buffer: fillParamBuf } }
      ]
    });

    {
      const dispatch = this._dispatch2D(Math.ceil(leftLen / WORKGROUP_SIZE));
      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.hashJoinFill);
      pass.setBindGroup(0, fillBindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);
    }

    const [leftIndex, rightIndex] = await Promise.all([
      this._readbackTypedArray(outLeftBuf, Uint32Array),
      this._readbackTypedArray(outRightBuf, Uint32Array)
    ]);

    leftKeysBuf.destroy();
    rightKeysBuf.destroy();
    tableKeysBuf.destroy();
    tableHeadsBuf.destroy();
    nextBuf.destroy();
    clearParamBuf.destroy();
    buildParamBuf.destroy();
    countParamBuf.destroy();
    fillParamBuf.destroy();
    offsetsBuf.destroy();
    countsBuf.destroy();
    outLeftBuf.destroy();
    outRightBuf.destroy();

    // Canonicalize output order for stable semantics across backends.
    if (total > 1) {
      const packed = new BigUint64Array(total);
      for (let i = 0; i < total; i++) {
        packed[i] = (BigInt(leftIndex[i]) << 32n) | BigInt(rightIndex[i]);
      }
      packed.sort();
      for (let i = 0; i < total; i++) {
        const v = packed[i];
        leftIndex[i] = Number(v >> 32n);
        rightIndex[i] = Number(v & 0xffff_ffffn);
      }
    }

    return { leftIndex, rightIndex };
  }

  /**
   * Explicit upload API: lets callers keep vectors resident on GPU to avoid
   * repeated CPU→GPU uploads across multiple kernels.
   * @param {Float32Array | Float64Array} values
   */
  uploadVector(values) {
    this._ensureNotDisposed();
    const buffer = this._createStorageBufferFromArray(values, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST | GPUBufferUsage.COPY_SRC);
    return new GpuVector(this, buffer, values.length, dtypeOf(values));
  }

  /**
   * @param {GpuVector} vec
   */
  async readbackVector(vec) {
    this._ensureNotDisposed();
    if (vec.dtype === "f32") return this._readbackTypedArray(vec.buffer, Float32Array);
    if (vec.dtype === "f64") return this._readbackTypedArray(vec.buffer, Float64Array);
    throw new Error(`Unsupported vector dtype: ${vec.dtype}`);
  }

  /**
   * @param {GpuVector} vec
   */
  async sumVector(vec) {
    this._precisionUsed.add(vec.dtype);
    const reduced = await this._reduce(vec.buffer, vec.length, vec.dtype, "sum");
    const out = await this._readbackTypedArray(reduced, vec.dtype === "f64" ? Float64Array : Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} vec
   */
  async minVector(vec) {
    this._precisionUsed.add(vec.dtype);
    const reduced = await this._reduce(vec.buffer, vec.length, vec.dtype, "min");
    const out = await this._readbackTypedArray(reduced, vec.dtype === "f64" ? Float64Array : Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} vec
   */
  async maxVector(vec) {
    this._precisionUsed.add(vec.dtype);
    const reduced = await this._reduce(vec.buffer, vec.length, vec.dtype, "max");
    const out = await this._readbackTypedArray(reduced, vec.dtype === "f64" ? Float64Array : Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} a
   * @param {GpuVector} b
   */
  async sumproductVectors(a, b) {
    if (a.dtype !== b.dtype) throw new Error(`SUMPRODUCT dtype mismatch: ${a.dtype} vs ${b.dtype}`);
    if (a.length !== b.length) throw new Error(`SUMPRODUCT length mismatch: ${a.length} vs ${b.length}`);

    this._precisionUsed.add(a.dtype);
    const reduced = await this._reduceSumproduct(a.buffer, b.buffer, a.length, a.dtype);
    const out = await this._readbackTypedArray(reduced, a.dtype === "f64" ? Float64Array : Float32Array);
    reduced.destroy();
    return out[0];
  }

  /**
   * @param {GPUBuffer} inputBuf
   * @param {number} length
   * @param {"f32" | "f64"} dtype
   * @returns {Promise<GPUBuffer>}
   */
  async _reduce(inputBuf, length, dtype, op) {
    let inBuf = inputBuf;
    let n = length;
    /** @type {GPUBuffer[]} */
    const temporaries = [];

    /** @type {GPUComputePipeline | undefined} */
    let pipeline;
    if (op === "sum") {
      pipeline = dtype === "f64" ? this.pipelines.reduceSum_f64 : this.pipelines.reduceSum_f32;
    } else if (op === "min") {
      pipeline = dtype === "f64" ? this.pipelines.reduceMin_f64 : this.pipelines.reduceMin_f32;
    } else if (op === "max") {
      pipeline = dtype === "f64" ? this.pipelines.reduceMax_f64 : this.pipelines.reduceMax_f32;
    }
    if (!pipeline) {
      throw new Error(`No ${dtype} reduction pipeline available for op=${op}`);
    }

    while (n > 1) {
      const requiredWorkgroups = Math.ceil(n / (WORKGROUP_SIZE * 2));
      const dispatch = this._dispatch2D(requiredWorkgroups);
      const outLen = dispatch.total;
      const outBuf = this.device.createBuffer({
        size: outLen * byteSizeOf(dtype),
        usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
      });

      const params = new Uint32Array([n, 0, 0, 0]);
      const paramBuf = this.device.createBuffer({
        size: 16,
        usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
      });
      this.queue.writeBuffer(paramBuf, 0, params);

      const bindGroup = this.device.createBindGroup({
        layout: pipeline.getBindGroupLayout(0),
        entries: [
          { binding: 0, resource: { buffer: inBuf } },
          { binding: 1, resource: { buffer: outBuf } },
          { binding: 2, resource: { buffer: paramBuf } }
        ]
      });

      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(pipeline);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
      pass.end();
      this.queue.submit([encoder.finish()]);

      if (inBuf !== inputBuf) temporaries.push(inBuf);
      temporaries.push(paramBuf);

      inBuf = outBuf;
      n = outLen;
    }

    // Ensure all queued reduction passes are complete before destroying buffers
    // used by submitted command buffers.
    await this.queue.onSubmittedWorkDone();
    for (const buf of temporaries) buf.destroy();
    return inBuf;
  }

  /**
   * @param {GPUBuffer} aBuf
   * @param {GPUBuffer} bBuf
   * @param {number} length
   * @param {"f32" | "f64"} dtype
   * @returns {Promise<GPUBuffer>}
   */
  async _reduceSumproduct(aBuf, bBuf, length, dtype) {
    let n = length;
    const requiredWorkgroups = Math.ceil(n / (WORKGROUP_SIZE * 2));
    const dispatch = this._dispatch2D(requiredWorkgroups);
    const outLen = dispatch.total;
    const outBuf = this.device.createBuffer({
      size: outLen * byteSizeOf(dtype),
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
    });

    const params = new Uint32Array([n, 0, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const pipeline = dtype === "f64" ? this.pipelines.reduceSumproduct_f64 : this.pipelines.reduceSumproduct_f32;
    if (!pipeline) {
      throw new Error(`No ${dtype} sumproduct pipeline available`);
    }

    const bindGroup = this.device.createBindGroup({
      layout: pipeline.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: aBuf } },
        { binding: 1, resource: { buffer: bBuf } },
        { binding: 2, resource: { buffer: outBuf } },
        { binding: 3, resource: { buffer: paramBuf } }
      ]
    });

    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(pipeline);
    pass.setBindGroup(0, bindGroup);
    pass.dispatchWorkgroups(dispatch.x, dispatch.y, 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    n = outLen;
    const reduced = await this._reduce(outBuf, n, dtype, "sum");
    paramBuf.destroy();
    if (reduced !== outBuf) outBuf.destroy();
    return reduced;
  }

  /**
   * @param {ArrayBufferView} view
   * @param {number} usage
   */
  _createStorageBufferFromArray(view, usage) {
    const buffer = this.device.createBuffer({
      size: view.byteLength,
      usage: usage | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(buffer, 0, view);
    return buffer;
  }

  /**
   * @template {TypedArrayConstructor} T
   * @param {GPUBuffer} src
   * @param {T} ctor
   * @returns {Promise<InstanceType<T>>}
   */
  async _readbackTypedArray(src, ctor) {
    const readBuf = this.device.createBuffer({
      size: src.size,
      usage: GPUBufferUsage.COPY_DST | GPUBufferUsage.MAP_READ
    });
    const encoder = this.device.createCommandEncoder();
    encoder.copyBufferToBuffer(src, 0, readBuf, 0, src.size);
    this.queue.submit([encoder.finish()]);

    await readBuf.mapAsync(GPUMapMode.READ);
    const copy = new ctor(readBuf.getMappedRange().slice(0));
    readBuf.unmap();
    readBuf.destroy();
    return copy;
  }
}

/**
 * @typedef {Float32ArrayConstructor | Float64ArrayConstructor | Uint32ArrayConstructor} TypedArrayConstructor
 */
