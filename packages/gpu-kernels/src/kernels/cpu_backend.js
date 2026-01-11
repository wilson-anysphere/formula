/**
 * @typedef {"cpu" | "webgpu"} BackendKind
 */

/**
 * CPU backend (fallback). Intended to be swapped for WASM/SIMD kernels in the
 * full engine; kept dependency-free here to make the WebGPU integration and
 * tests self-contained.
 */
export class CpuBackend {
  /** @type {BackendKind} */
  kind = "cpu";

  /**
   * @returns {{ kind: BackendKind, supportedKernels: Record<string, boolean> }}
   */
  diagnostics() {
    return {
      kind: this.kind,
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
      }
    };
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async sum(values) {
    let acc = 0;
    for (let i = 0; i < values.length; i++) acc += values[i];
    return acc;
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async min(values) {
    if (values.length === 0) return Number.POSITIVE_INFINITY;
    let acc = Number.POSITIVE_INFINITY;
    for (let i = 0; i < values.length; i++) acc = Math.min(acc, values[i]);
    return acc;
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async max(values) {
    if (values.length === 0) return Number.NEGATIVE_INFINITY;
    let acc = Number.NEGATIVE_INFINITY;
    for (let i = 0; i < values.length; i++) acc = Math.max(acc, values[i]);
    return acc;
  }

  /**
   * @param {Float32Array | Float64Array} values
   */
  async average(values) {
    if (values.length === 0) return Number.NaN;
    let acc = 0;
    for (let i = 0; i < values.length; i++) acc += values[i];
    return acc / values.length;
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
   */
  async sumproduct(a, b) {
    if (a.length !== b.length) {
      throw new Error(`SUMPRODUCT length mismatch: ${a.length} vs ${b.length}`);
    }
    let acc = 0;
    for (let i = 0; i < a.length; i++) acc += a[i] * b[i];
    return acc;
  }

  /**
   * Matrix multiply (row-major).
   * @param {Float32Array | Float64Array} a
   * @param {Float32Array | Float64Array} b
   * @param {number} aRows
   * @param {number} aCols
   * @param {number} bCols
   * @returns {Promise<Float64Array>}
   */
  async mmult(a, b, aRows, aCols, bCols) {
    if (a.length !== aRows * aCols) {
      throw new Error(`MMULT A shape mismatch: a.length=${a.length} vs ${aRows}x${aCols}`);
    }
    if (b.length !== aCols * bCols) {
      throw new Error(`MMULT B shape mismatch: b.length=${b.length} vs ${aCols}x${bCols}`);
    }

    const out = new Float64Array(aRows * bCols);
    for (let r = 0; r < aRows; r++) {
      for (let c = 0; c < bCols; c++) {
        let acc = 0;
        const aRow = r * aCols;
        for (let k = 0; k < aCols; k++) {
          acc += a[aRow + k] * b[k * bCols + c];
        }
        out[r * bCols + c] = acc;
      }
    }
    return out;
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<Float64Array>}
   */
  async sort(values) {
    const out = Float64Array.from(values);
    out.sort();
    return out;
  }

  /**
   * @param {Float32Array | Float64Array} values
   * @param {{ min: number, max: number, bins: number }} opts
   * @returns {Promise<Uint32Array>}
   */
  async histogram(values, opts) {
    const { min, max, bins } = opts;
    if (!(bins > 0)) throw new Error("histogram bins must be > 0");
    if (!(max > min)) throw new Error("histogram max must be > min");
    const counts = new Uint32Array(bins);
    const invWidth = bins / (max - min);

    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      if (Number.isNaN(v)) continue;
      let bin;
      if (v <= min) {
        bin = 0;
      } else if (v >= max) {
        bin = bins - 1;
      } else {
        bin = Math.floor((v - min) * invWidth);
      }
      if (bin < 0) bin = 0;
      if (bin >= bins) bin = bins - 1;
      counts[bin] += 1;
    }

    return counts;
  }
}
