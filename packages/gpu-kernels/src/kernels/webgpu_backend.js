import { loadTextResource } from "../util/load_text_resource.js";

/**
 * @typedef {"webgpu"} BackendKind
 */

const WORKGROUP_SIZE = 256;
const MMULT_TILE = 8;

function hasWebGpu() {
  return typeof navigator !== "undefined" && navigator && "gpu" in navigator && navigator.gpu;
}

function nextPowerOfTwo(n) {
  let p = 1;
  while (p < n) p <<= 1;
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

      const [reduceSumSrc, reduceMinSrc, reduceMaxSrc, reduceSumproductSrc, mmultSrc, histogramSrc, bitonicSortSrc] = await Promise.all([
        loadTextResource(new URL("./wgsl/reduce_sum.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_min.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_max.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_sumproduct.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/mmult.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/histogram.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/bitonic_sort.wgsl", import.meta.url))
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

      /** @type {WebGpuBackend["pipelines"]} */
      const pipelines = {
        reduceSum_f32,
        reduceMin_f32,
        reduceMax_f32,
        reduceSumproduct_f32,
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
        mmult: true,
        sort: true,
        histogram: true
      },
      supportedKernelsF64: {
        sum: Boolean(this.pipelines.reduceSum_f64),
        min: Boolean(this.pipelines.reduceMin_f64),
        max: Boolean(this.pipelines.reduceMax_f64),
        sumproduct: Boolean(this.pipelines.reduceSumproduct_f64),
        average: Boolean(this.pipelines.reduceSum_f64),
        count: Boolean(this.pipelines.reduceSum_f64),
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
   * @param {"sum" | "min" | "max" | "average" | "count" | "sumproduct" | "mmult" | "sort" | "histogram"} kernel
   * @param {"f32" | "f64"} dtype
   */
  supportsKernelPrecision(kernel, dtype) {
    if (dtype === "f32") return true;
    if (!this.supportsF64) return false;
    switch (kernel) {
      case "sum":
        return Boolean(this.pipelines.reduceSum_f64);
      case "min":
        return Boolean(this.pipelines.reduceMin_f64);
      case "max":
        return Boolean(this.pipelines.reduceMax_f64);
      case "sumproduct":
        return Boolean(this.pipelines.reduceSumproduct_f64);
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
    pass.dispatchWorkgroups(Math.ceil(bCols / MMULT_TILE), Math.ceil(aRows / MMULT_TILE), 1);
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
    if (padded !== n) data.fill(Number.POSITIVE_INFINITY, n);

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
    const reduced = await this._reduce(vec.buffer, vec.length, vec.dtype, "sum");
    const out = await this._readbackTypedArray(reduced, vec.dtype === "f64" ? Float64Array : Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} vec
   */
  async minVector(vec) {
    const reduced = await this._reduce(vec.buffer, vec.length, vec.dtype, "min");
    const out = await this._readbackTypedArray(reduced, vec.dtype === "f64" ? Float64Array : Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} vec
   */
  async maxVector(vec) {
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
