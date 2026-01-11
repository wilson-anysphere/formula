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
        groupByCount: true,
        groupBySum: true,
        groupByMin: true,
        groupByMax: true,
        hashJoin: true,
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
        if (!Number.isFinite(bin)) bin = bin === Number.POSITIVE_INFINITY ? bins - 1 : 0;
      }
      if (bin < 0) bin = 0;
      if (bin >= bins) bin = bins - 1;
      counts[bin] += 1;
    }

    return counts;
  }

  /**
   * Group-by COUNT.
   * Keys may be `Uint32Array` (dictionary ids) or `Int32Array` (signed keys).
   * Returned keys are sorted ascending by the key type's numeric ordering.
   *
   * @param {Uint32Array | Int32Array} keys
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, counts: Uint32Array }>}
   */
  async groupByCount(keys) {
    /** @type {Map<number, number>} */
    const countsMap = new Map();
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      countsMap.set(k, (countsMap.get(k) ?? 0) + 1);
    }

    const unique = Array.from(countsMap.keys());
    unique.sort((a, b) => a - b);

    const uniqueKeys = keys instanceof Int32Array ? new Int32Array(unique.length) : new Uint32Array(unique.length);
    const counts = new Uint32Array(unique.length);
    for (let i = 0; i < unique.length; i++) {
      const k = unique[i];
      uniqueKeys[i] = k;
      counts[i] = countsMap.get(k) ?? 0;
    }
    return { uniqueKeys, counts };
  }

  /**
   * Group-by SUM(+COUNT). SUM follows JS numeric semantics: `NaN` propagates and
   * `±Infinity` behaves per IEEE-754.
   *
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, sums: Float64Array, counts: Uint32Array }>}
   */
  async groupBySum(keys, values) {
    if (keys.length !== values.length) {
      throw new Error(`groupBySum length mismatch: keys=${keys.length} values=${values.length}`);
    }

    /** @type {Map<number, { sum: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const v = values[i];
      const entry = map.get(k);
      if (entry) {
        entry.sum += v;
        entry.count += 1;
      } else {
        map.set(k, { sum: v, count: 1 });
      }
    }

    const unique = Array.from(map.keys());
    unique.sort((a, b) => a - b);

    const uniqueKeys = keys instanceof Int32Array ? new Int32Array(unique.length) : new Uint32Array(unique.length);
    const sums = new Float64Array(unique.length);
    const counts = new Uint32Array(unique.length);
    for (let i = 0; i < unique.length; i++) {
      const k = unique[i];
      const entry = map.get(k);
      uniqueKeys[i] = k;
      sums[i] = entry?.sum ?? 0;
      counts[i] = entry?.count ?? 0;
    }
    return { uniqueKeys, sums, counts };
  }

  /**
   * Group-by MIN(+COUNT). MIN follows JS `Math.min` semantics: `NaN` propagates
   * and signed zero is preserved (`min(0, -0) === -0`).
   *
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, mins: Float64Array, counts: Uint32Array }>}
   */
  async groupByMin(keys, values) {
    if (keys.length !== values.length) {
      throw new Error(`groupByMin length mismatch: keys=${keys.length} values=${values.length}`);
    }

    /** @type {Map<number, { min: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const v = values[i];
      const entry = map.get(k);
      if (entry) {
        entry.min = Math.min(entry.min, v);
        entry.count += 1;
      } else {
        map.set(k, { min: v, count: 1 });
      }
    }

    const unique = Array.from(map.keys());
    unique.sort((a, b) => a - b);

    const uniqueKeys = keys instanceof Int32Array ? new Int32Array(unique.length) : new Uint32Array(unique.length);
    const mins = new Float64Array(unique.length);
    const counts = new Uint32Array(unique.length);
    for (let i = 0; i < unique.length; i++) {
      const k = unique[i];
      const entry = map.get(k);
      uniqueKeys[i] = k;
      mins[i] = entry?.min ?? Number.POSITIVE_INFINITY;
      counts[i] = entry?.count ?? 0;
    }
    return { uniqueKeys, mins, counts };
  }

  /**
   * Group-by MAX(+COUNT). MAX follows JS `Math.max` semantics: `NaN` propagates
   * and signed zero is preserved (`max(0, -0) === 0`).
   *
   * @param {Uint32Array | Int32Array} keys
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeys: Uint32Array | Int32Array, maxs: Float64Array, counts: Uint32Array }>}
   */
  async groupByMax(keys, values) {
    if (keys.length !== values.length) {
      throw new Error(`groupByMax length mismatch: keys=${keys.length} values=${values.length}`);
    }

    /** @type {Map<number, { max: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const v = values[i];
      const entry = map.get(k);
      if (entry) {
        entry.max = Math.max(entry.max, v);
        entry.count += 1;
      } else {
        map.set(k, { max: v, count: 1 });
      }
    }

    const unique = Array.from(map.keys());
    unique.sort((a, b) => a - b);

    const uniqueKeys = keys instanceof Int32Array ? new Int32Array(unique.length) : new Uint32Array(unique.length);
    const maxs = new Float64Array(unique.length);
    const counts = new Uint32Array(unique.length);
    for (let i = 0; i < unique.length; i++) {
      const k = unique[i];
      const entry = map.get(k);
      uniqueKeys[i] = k;
      maxs[i] = entry?.max ?? Number.NEGATIVE_INFINITY;
      counts[i] = entry?.count ?? 0;
    }
    return { uniqueKeys, maxs, counts };
  }

  /**
   * Inner hash-join of two key arrays. Returns pairs of matching indices.
   *
   * Output pairs are sorted by `(leftIndex, rightIndex)` ascending.
   *
   * @param {Uint32Array | Int32Array} leftKeys
   * @param {Uint32Array | Int32Array} rightKeys
   * @param {{ joinType?: "inner" | "left" }} [opts]
   * @returns {Promise<{ leftIndex: Uint32Array, rightIndex: Uint32Array }>}
   */
  async hashJoin(leftKeys, rightKeys, opts = {}) {
    const joinType = opts.joinType ?? "inner";
    if (joinType !== "inner" && joinType !== "left") {
      throw new Error(`hashJoin joinType must be "inner" | "left", got ${String(joinType)}`);
    }

    if (leftKeys.length === 0) {
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }
    if (rightKeys.length === 0) {
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
    const leftSigned = leftKeys instanceof Int32Array;
    const rightSigned = rightKeys instanceof Int32Array;
    if (leftSigned !== rightSigned) {
      throw new Error(
        `hashJoin key type mismatch: left=${leftSigned ? "i32" : "u32"} right=${rightSigned ? "i32" : "u32"} (pass matching Int32Array/Uint32Array types)`
      );
    }

    /** @type {Map<number, number[]>} */
    const rightMap = new Map();
    for (let j = 0; j < rightKeys.length; j++) {
      const k = rightKeys[j];
      const arr = rightMap.get(k);
      if (arr) arr.push(j);
      else rightMap.set(k, [j]);
    }

    let total = 0;
    for (let i = 0; i < leftKeys.length; i++) {
      const arr = rightMap.get(leftKeys[i]);
      if (arr) total += arr.length;
      else if (joinType === "left") total += 1;
    }

    const leftIndex = new Uint32Array(total);
    const rightIndex = new Uint32Array(total);

    let p = 0;
    for (let i = 0; i < leftKeys.length; i++) {
      const arr = rightMap.get(leftKeys[i]);
      if (arr) {
        for (let k = 0; k < arr.length; k++) {
          leftIndex[p] = i;
          rightIndex[p] = arr[k];
          p += 1;
        }
      } else if (joinType === "left") {
        leftIndex[p] = i;
        rightIndex[p] = 0xffff_ffff;
        p += 1;
      }
    }

    return { leftIndex, rightIndex };
  }
}
