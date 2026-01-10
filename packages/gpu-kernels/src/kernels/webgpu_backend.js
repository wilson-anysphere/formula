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
   *  pipelines: {
   *    reduceSum: GPUComputePipeline,
   *    reduceSumproduct: GPUComputePipeline,
   *    mmult: GPUComputePipeline,
   *    histogram: GPUComputePipeline,
   *    bitonicSort: GPUComputePipeline
   *  }
   * }} resources
   */
  constructor(device, adapter, resources) {
    this.device = device;
    this.adapter = adapter;
    this.queue = device.queue;
    this.pipelines = resources.pipelines;
    this._disposed = false;
  }

  static async createIfSupported() {
    try {
      if (!hasWebGpu()) return null;

      /** @type {GPUAdapter | null} */
      const adapter = await navigator.gpu.requestAdapter({ powerPreference: "high-performance" });
      if (!adapter) return null;

      /** @type {GPUDevice} */
      const device = await adapter.requestDevice();

      const [reduceSumSrc, reduceSumproductSrc, mmultSrc, histogramSrc, bitonicSortSrc] = await Promise.all([
        loadTextResource(new URL("./wgsl/reduce_sum.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/reduce_sumproduct.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/mmult.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/histogram.wgsl", import.meta.url)),
        loadTextResource(new URL("./wgsl/bitonic_sort.wgsl", import.meta.url))
      ]);

      const reduceSum = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceSumSrc }),
          entryPoint: "main"
        }
      });

      const reduceSumproduct = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: reduceSumproductSrc }),
          entryPoint: "main"
        }
      });

      const mmult = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: mmultSrc }),
          entryPoint: "main"
        }
      });

      const histogram = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: histogramSrc }),
          entryPoint: "main"
        }
      });

      const bitonicSort = device.createComputePipeline({
        layout: "auto",
        compute: {
          module: device.createShaderModule({ code: bitonicSortSrc }),
          entryPoint: "main"
        }
      });

      return new WebGpuBackend(device, adapter, {
        pipelines: { reduceSum, reduceSumproduct, mmult, histogram, bitonicSort }
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
        sumproduct: true,
        mmult: true,
        sort: true,
        histogram: true
      },
      numericPrecision: "f32",
      workgroupSize: WORKGROUP_SIZE
    };
  }

  _ensureNotDisposed() {
    if (this._disposed) throw new Error("WebGpuBackend is disposed");
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64: boolean }} opts
   */
  async sum(values, opts) {
    this._ensureNotDisposed();
    const dtype = dtypeOf(values);
    if (dtype === "f64" && !opts.allowFp32FallbackForF64) {
      throw new Error("WebGPU backend only supports f32 kernels; disallowing f64->f32 fallback");
    }
    const f32 = toFloat32(values);
    if (f32.length === 0) return 0;
    const vec = this.uploadVector(f32);
    try {
      return await this.sumVector(vec);
    } finally {
      vec.destroy();
    }
  }

  /**
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   * @param {{ allowFp32FallbackForF64: boolean }} opts
   */
  async sumproduct(a, b, opts) {
    this._ensureNotDisposed();
    if (a.length !== b.length) throw new Error(`SUMPRODUCT length mismatch: ${a.length} vs ${b.length}`);

    const aDtype = dtypeOf(a);
    const bDtype = dtypeOf(b);
    if ((aDtype === "f64" || bDtype === "f64") && !opts.allowFp32FallbackForF64) {
      throw new Error("WebGPU backend only supports f32 kernels; disallowing f64->f32 fallback");
    }

    const a32 = toFloat32(a);
    const b32 = toFloat32(b);
    if (a32.length === 0) return 0;
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
   * @param {{ allowFp32FallbackForF64: boolean }} opts
   */
  async mmult(a, b, aRows, aCols, bCols, opts) {
    this._ensureNotDisposed();
    const aDtype = dtypeOf(a);
    const bDtype = dtypeOf(b);
    if ((aDtype === "f64" || bDtype === "f64") && !opts.allowFp32FallbackForF64) {
      throw new Error("WebGPU backend only supports f32 kernels; disallowing f64->f32 fallback");
    }

    const a32 = toFloat32(a);
    const b32 = toFloat32(b);

    const aBuf = this._createStorageBufferFromArray(a32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const bBuf = this._createStorageBufferFromArray(b32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);
    const outBytes = aRows * bCols * 4;
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
      layout: this.pipelines.mmult.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: aBuf } },
        { binding: 1, resource: { buffer: bBuf } },
        { binding: 2, resource: { buffer: outBuf } },
        { binding: 3, resource: { buffer: paramBuf } }
      ]
    });

    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(this.pipelines.mmult);
    pass.setBindGroup(0, bindGroup);
    pass.dispatchWorkgroups(Math.ceil(bCols / MMULT_TILE), Math.ceil(aRows / MMULT_TILE), 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    const out = await this._readbackTypedArray(outBuf, Float32Array);

    aBuf.destroy();
    bBuf.destroy();
    outBuf.destroy();
    paramBuf.destroy();

    // Promote to f64 for API consistency.
    return Float64Array.from(out);
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ allowFp32FallbackForF64: boolean }} opts
   */
  async sort(values, opts) {
    this._ensureNotDisposed();
    const dtype = dtypeOf(values);
    if (dtype === "f64" && !opts.allowFp32FallbackForF64) {
      throw new Error("WebGPU backend only supports f32 kernels; disallowing f64->f32 fallback");
    }
    const f32 = toFloat32(values);
    const n = f32.length;
    if (n === 0) return new Float64Array();
    const padded = nextPowerOfTwo(n);
    const data = new Float32Array(padded);
    data.set(f32);
    if (padded !== n) data.fill(Number.POSITIVE_INFINITY, n);

    const dataBuf = this._createStorageBufferFromArray(data, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC);

    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.bitonicSort.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: dataBuf } },
        { binding: 1, resource: { buffer: paramBuf } }
      ]
    });

    for (let k = 2; k <= padded; k <<= 1) {
      for (let j = k >> 1; j > 0; j >>= 1) {
        const params = new Uint32Array([padded, j, k, 0]);
        this.queue.writeBuffer(paramBuf, 0, params);

        const encoder = this.device.createCommandEncoder();
        const pass = encoder.beginComputePass();
        pass.setPipeline(this.pipelines.bitonicSort);
        pass.setBindGroup(0, bindGroup);
        pass.dispatchWorkgroups(Math.ceil(padded / WORKGROUP_SIZE), 1, 1);
        pass.end();
        this.queue.submit([encoder.finish()]);
      }
    }

    const outPadded = await this._readbackTypedArray(dataBuf, Float32Array);
    dataBuf.destroy();
    paramBuf.destroy();

    return Float64Array.from(outPadded.subarray(0, n));
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ min: number, max: number, bins: number }} opts
   * @param {{ allowFp32FallbackForF64: boolean }} backendOpts
   */
  async histogram(values, opts, backendOpts) {
    this._ensureNotDisposed();
    const dtype = dtypeOf(values);
    if (dtype === "f64" && !backendOpts.allowFp32FallbackForF64) {
      throw new Error("WebGPU backend only supports f32 kernels; disallowing f64->f32 fallback");
    }
    const f32 = toFloat32(values);
    const { min, max, bins } = opts;
    if (!(bins > 0)) throw new Error("histogram bins must be > 0");
    if (!(max > min)) throw new Error("histogram max must be > min");

    if (f32.length === 0) {
      return new Uint32Array(bins);
    }

    const inputBuf = this._createStorageBufferFromArray(f32, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST);

    const binBuf = this.device.createBuffer({
      size: bins * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(binBuf, 0, new Uint32Array(bins));

    const invBinWidth = bins / (max - min);
    const params = new Uint32Array(8);
    new Float32Array(params.buffer, 0, 3).set([min, max, invBinWidth]);
    params[3] = f32.length;
    params[4] = bins;
    // params[5..7] padding

    const paramBuf = this.device.createBuffer({
      size: 32,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.histogram.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: inputBuf } },
        { binding: 1, resource: { buffer: binBuf } },
        { binding: 2, resource: { buffer: paramBuf } }
      ]
    });

    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(this.pipelines.histogram);
    pass.setBindGroup(0, bindGroup);
    pass.dispatchWorkgroups(Math.ceil(f32.length / WORKGROUP_SIZE), 1, 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    const counts = await this._readbackTypedArray(binBuf, Uint32Array);
    inputBuf.destroy();
    binBuf.destroy();
    paramBuf.destroy();
    return counts;
  }

  /**
   * Explicit upload API: lets callers keep vectors resident on GPU to avoid
   * repeated CPU→GPU uploads across multiple kernels.
   * @param {Float32Array} values
   */
  uploadVector(values) {
    this._ensureNotDisposed();
    const buffer = this._createStorageBufferFromArray(values, GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_DST | GPUBufferUsage.COPY_SRC);
    return new GpuVector(this, buffer, values.length, "f32");
  }

  /**
   * @param {GpuVector} vec
   */
  async readbackVector(vec) {
    this._ensureNotDisposed();
    if (vec.dtype !== "f32") throw new Error(`readbackVector only supports f32`);
    return this._readbackTypedArray(vec.buffer, Float32Array);
  }

  /**
   * @param {GpuVector} vec
   */
  async sumVector(vec) {
    if (vec.dtype !== "f32") throw new Error(`sumVector only supports f32`);
    const reduced = await this._reduce(vec.buffer, vec.length);
    const out = await this._readbackTypedArray(reduced, Float32Array);
    if (reduced !== vec.buffer) reduced.destroy();
    return out[0];
  }

  /**
   * @param {GpuVector} a
   * @param {GpuVector} b
   */
  async sumproductVectors(a, b) {
    if (a.dtype !== "f32" || b.dtype !== "f32") throw new Error(`sumproductVectors only supports f32`);
    if (a.length !== b.length) throw new Error(`SUMPRODUCT length mismatch: ${a.length} vs ${b.length}`);

    const reduced = await this._reduceSumproduct(a.buffer, b.buffer, a.length);
    const out = await this._readbackTypedArray(reduced, Float32Array);
    reduced.destroy();
    return out[0];
  }

  /**
   * @param {GPUBuffer} inputBuf
   * @param {number} length
   * @returns {Promise<GPUBuffer>}
   */
  async _reduce(inputBuf, length) {
    let inBuf = inputBuf;
    let n = length;
    /** @type {GPUBuffer[]} */
    const temporaries = [];

    while (n > 1) {
      const outLen = Math.ceil(n / (WORKGROUP_SIZE * 2));
      const outBuf = this.device.createBuffer({
        size: outLen * 4,
        usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
      });

      const params = new Uint32Array([n, 0, 0, 0]);
      const paramBuf = this.device.createBuffer({
        size: 16,
        usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
      });
      this.queue.writeBuffer(paramBuf, 0, params);

      const bindGroup = this.device.createBindGroup({
        layout: this.pipelines.reduceSum.getBindGroupLayout(0),
        entries: [
          { binding: 0, resource: { buffer: inBuf } },
          { binding: 1, resource: { buffer: outBuf } },
          { binding: 2, resource: { buffer: paramBuf } }
        ]
      });

      const encoder = this.device.createCommandEncoder();
      const pass = encoder.beginComputePass();
      pass.setPipeline(this.pipelines.reduceSum);
      pass.setBindGroup(0, bindGroup);
      pass.dispatchWorkgroups(outLen, 1, 1);
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
   * @returns {Promise<GPUBuffer>}
   */
  async _reduceSumproduct(aBuf, bBuf, length) {
    let n = length;
    const outLen = Math.ceil(n / (WORKGROUP_SIZE * 2));
    const outBuf = this.device.createBuffer({
      size: outLen * 4,
      usage: GPUBufferUsage.STORAGE | GPUBufferUsage.COPY_SRC
    });

    const params = new Uint32Array([n, 0, 0, 0]);
    const paramBuf = this.device.createBuffer({
      size: 16,
      usage: GPUBufferUsage.UNIFORM | GPUBufferUsage.COPY_DST
    });
    this.queue.writeBuffer(paramBuf, 0, params);

    const bindGroup = this.device.createBindGroup({
      layout: this.pipelines.reduceSumproduct.getBindGroupLayout(0),
      entries: [
        { binding: 0, resource: { buffer: aBuf } },
        { binding: 1, resource: { buffer: bBuf } },
        { binding: 2, resource: { buffer: outBuf } },
        { binding: 3, resource: { buffer: paramBuf } }
      ]
    });

    const encoder = this.device.createCommandEncoder();
    const pass = encoder.beginComputePass();
    pass.setPipeline(this.pipelines.reduceSumproduct);
    pass.setBindGroup(0, bindGroup);
    pass.dispatchWorkgroups(outLen, 1, 1);
    pass.end();
    this.queue.submit([encoder.finish()]);

    n = outLen;
    const reduced = await this._reduce(outBuf, n);
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
