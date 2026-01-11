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
        groupByCount2: true,
        groupBySum2: true,
        groupByMin2: true,
        groupByMax2: true,
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
   * Group-by COUNT on two key columns.
   *
   * Keys may be `Uint32Array` (dictionary ids) or `Int32Array` (signed keys).
   * Returned keys are sorted lexicographically by `(keyA, keyB)` using each
   * key column's numeric ordering.
   *
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @returns {Promise<{ uniqueKeysA: Uint32Array | Int32Array, uniqueKeysB: Uint32Array | Int32Array, counts: Uint32Array }>}
   */
  async groupByCount2(keysA, keysB) {
    if (keysA.length !== keysB.length) {
      throw new Error(`groupByCount2 length mismatch: keysA=${keysA.length} keysB=${keysB.length}`);
    }
    const n = keysA.length;
    if (n === 0) {
      return {
        uniqueKeysA: keysA instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        uniqueKeysB: keysB instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        counts: new Uint32Array()
      };
    }

    const signedA = keysA instanceof Int32Array;
    const signedB = keysB instanceof Int32Array;

    /** @type {Map<bigint, number>} */
    const countsMap = new Map();
    for (let i = 0; i < n; i++) {
      const aBits = keysA[i] >>> 0;
      const bBits = keysB[i] >>> 0;
      const key = (BigInt(aBits) << 32n) | BigInt(bBits);
      countsMap.set(key, (countsMap.get(key) ?? 0) + 1);
    }

    /** @type {{ key: bigint, sortKey: bigint, count: number }[]} */
    const entries = [];
    for (const [key, count] of countsMap) {
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      const aSort = (aBits ^ (signedA ? 0x8000_0000 : 0)) >>> 0;
      const bSort = (bBits ^ (signedB ? 0x8000_0000 : 0)) >>> 0;
      const sortKey = (BigInt(aSort) << 32n) | BigInt(bSort);
      entries.push({ key, sortKey, count });
    }
    entries.sort((a, b) => (a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0));

    const uniqueKeysA = signedA ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const uniqueKeysB = signedB ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const counts = new Uint32Array(entries.length);
    for (let i = 0; i < entries.length; i++) {
      const key = entries[i].key;
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      uniqueKeysA[i] = signedA ? aBits | 0 : aBits;
      uniqueKeysB[i] = signedB ? bBits | 0 : bBits;
      counts[i] = entries[i].count;
    }

    return { uniqueKeysA, uniqueKeysB, counts };
  }

  /**
   * Group-by SUM(+COUNT) on two key columns.
   *
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeysA: Uint32Array | Int32Array, uniqueKeysB: Uint32Array | Int32Array, sums: Float64Array, counts: Uint32Array }>}
   */
  async groupBySum2(keysA, keysB, values) {
    if (keysA.length !== keysB.length) {
      throw new Error(`groupBySum2 length mismatch: keysA=${keysA.length} keysB=${keysB.length}`);
    }
    if (keysA.length !== values.length) {
      throw new Error(`groupBySum2 length mismatch: keys=${keysA.length} values=${values.length}`);
    }
    const n = keysA.length;
    if (n === 0) {
      return {
        uniqueKeysA: keysA instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        uniqueKeysB: keysB instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        sums: new Float64Array(),
        counts: new Uint32Array()
      };
    }

    const signedA = keysA instanceof Int32Array;
    const signedB = keysB instanceof Int32Array;

    /** @type {Map<bigint, { sum: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < n; i++) {
      const aBits = keysA[i] >>> 0;
      const bBits = keysB[i] >>> 0;
      const key = (BigInt(aBits) << 32n) | BigInt(bBits);
      const entry = map.get(key);
      if (entry) {
        entry.sum += values[i];
        entry.count += 1;
      } else {
        map.set(key, { sum: values[i], count: 1 });
      }
    }

    /** @type {{ key: bigint, sortKey: bigint, sum: number, count: number }[]} */
    const entries = [];
    for (const [key, entry] of map) {
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      const aSort = (aBits ^ (signedA ? 0x8000_0000 : 0)) >>> 0;
      const bSort = (bBits ^ (signedB ? 0x8000_0000 : 0)) >>> 0;
      const sortKey = (BigInt(aSort) << 32n) | BigInt(bSort);
      entries.push({ key, sortKey, sum: entry.sum, count: entry.count });
    }
    entries.sort((a, b) => (a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0));

    const uniqueKeysA = signedA ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const uniqueKeysB = signedB ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const sums = new Float64Array(entries.length);
    const counts = new Uint32Array(entries.length);
    for (let i = 0; i < entries.length; i++) {
      const key = entries[i].key;
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      uniqueKeysA[i] = signedA ? aBits | 0 : aBits;
      uniqueKeysB[i] = signedB ? bBits | 0 : bBits;
      sums[i] = entries[i].sum;
      counts[i] = entries[i].count;
    }

    return { uniqueKeysA, uniqueKeysB, sums, counts };
  }

  /**
   * Group-by MIN(+COUNT) on two key columns.
   *
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeysA: Uint32Array | Int32Array, uniqueKeysB: Uint32Array | Int32Array, mins: Float64Array, counts: Uint32Array }>}
   */
  async groupByMin2(keysA, keysB, values) {
    if (keysA.length !== keysB.length) {
      throw new Error(`groupByMin2 length mismatch: keysA=${keysA.length} keysB=${keysB.length}`);
    }
    if (keysA.length !== values.length) {
      throw new Error(`groupByMin2 length mismatch: keys=${keysA.length} values=${values.length}`);
    }
    const n = keysA.length;
    if (n === 0) {
      return {
        uniqueKeysA: keysA instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        uniqueKeysB: keysB instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        mins: new Float64Array(),
        counts: new Uint32Array()
      };
    }

    const signedA = keysA instanceof Int32Array;
    const signedB = keysB instanceof Int32Array;

    /** @type {Map<bigint, { min: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < n; i++) {
      const aBits = keysA[i] >>> 0;
      const bBits = keysB[i] >>> 0;
      const key = (BigInt(aBits) << 32n) | BigInt(bBits);
      const entry = map.get(key);
      if (entry) {
        entry.min = Math.min(entry.min, values[i]);
        entry.count += 1;
      } else {
        map.set(key, { min: values[i], count: 1 });
      }
    }

    /** @type {{ key: bigint, sortKey: bigint, min: number, count: number }[]} */
    const entries = [];
    for (const [key, entry] of map) {
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      const aSort = (aBits ^ (signedA ? 0x8000_0000 : 0)) >>> 0;
      const bSort = (bBits ^ (signedB ? 0x8000_0000 : 0)) >>> 0;
      const sortKey = (BigInt(aSort) << 32n) | BigInt(bSort);
      entries.push({ key, sortKey, min: entry.min, count: entry.count });
    }
    entries.sort((a, b) => (a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0));

    const uniqueKeysA = signedA ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const uniqueKeysB = signedB ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const mins = new Float64Array(entries.length);
    const counts = new Uint32Array(entries.length);
    for (let i = 0; i < entries.length; i++) {
      const key = entries[i].key;
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      uniqueKeysA[i] = signedA ? aBits | 0 : aBits;
      uniqueKeysB[i] = signedB ? bBits | 0 : bBits;
      mins[i] = entries[i].min;
      counts[i] = entries[i].count;
    }

    return { uniqueKeysA, uniqueKeysB, mins, counts };
  }

  /**
   * Group-by MAX(+COUNT) on two key columns.
   *
   * @param {Uint32Array | Int32Array} keysA
   * @param {Uint32Array | Int32Array} keysB
   * @param {Float32Array | Float64Array} values
   * @returns {Promise<{ uniqueKeysA: Uint32Array | Int32Array, uniqueKeysB: Uint32Array | Int32Array, maxs: Float64Array, counts: Uint32Array }>}
   */
  async groupByMax2(keysA, keysB, values) {
    if (keysA.length !== keysB.length) {
      throw new Error(`groupByMax2 length mismatch: keysA=${keysA.length} keysB=${keysB.length}`);
    }
    if (keysA.length !== values.length) {
      throw new Error(`groupByMax2 length mismatch: keys=${keysA.length} values=${values.length}`);
    }
    const n = keysA.length;
    if (n === 0) {
      return {
        uniqueKeysA: keysA instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        uniqueKeysB: keysB instanceof Int32Array ? new Int32Array() : new Uint32Array(),
        maxs: new Float64Array(),
        counts: new Uint32Array()
      };
    }

    const signedA = keysA instanceof Int32Array;
    const signedB = keysB instanceof Int32Array;

    /** @type {Map<bigint, { max: number, count: number }>} */
    const map = new Map();
    for (let i = 0; i < n; i++) {
      const aBits = keysA[i] >>> 0;
      const bBits = keysB[i] >>> 0;
      const key = (BigInt(aBits) << 32n) | BigInt(bBits);
      const entry = map.get(key);
      if (entry) {
        entry.max = Math.max(entry.max, values[i]);
        entry.count += 1;
      } else {
        map.set(key, { max: values[i], count: 1 });
      }
    }

    /** @type {{ key: bigint, sortKey: bigint, max: number, count: number }[]} */
    const entries = [];
    for (const [key, entry] of map) {
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      const aSort = (aBits ^ (signedA ? 0x8000_0000 : 0)) >>> 0;
      const bSort = (bBits ^ (signedB ? 0x8000_0000 : 0)) >>> 0;
      const sortKey = (BigInt(aSort) << 32n) | BigInt(bSort);
      entries.push({ key, sortKey, max: entry.max, count: entry.count });
    }
    entries.sort((a, b) => (a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0));

    const uniqueKeysA = signedA ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const uniqueKeysB = signedB ? new Int32Array(entries.length) : new Uint32Array(entries.length);
    const maxs = new Float64Array(entries.length);
    const counts = new Uint32Array(entries.length);
    for (let i = 0; i < entries.length; i++) {
      const key = entries[i].key;
      const aBits = Number(key >> 32n) >>> 0;
      const bBits = Number(key & 0xffff_ffffn) >>> 0;
      uniqueKeysA[i] = signedA ? aBits | 0 : aBits;
      uniqueKeysB[i] = signedB ? bBits | 0 : bBits;
      maxs[i] = entries[i].max;
      counts[i] = entries[i].count;
    }

    return { uniqueKeysA, uniqueKeysB, maxs, counts };
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

    const EMPTY_U32 = 0xffff_ffff;
    const EMPTY_KEY = 0xffff_ffff;

    const leftU32 = leftKeys instanceof Uint32Array ? leftKeys : new Uint32Array(leftKeys.buffer, leftKeys.byteOffset, leftKeys.length);
    const rightU32 = rightKeys instanceof Uint32Array ? rightKeys : new Uint32Array(rightKeys.buffer, rightKeys.byteOffset, rightKeys.length);
    const leftLen = leftU32.length;
    const rightLen = rightU32.length;

    function nextPowerOfTwo(n) {
      let p = 1;
      while (p < n) p <<= 1;
      return p;
    }

    // Open-addressing hash table from rightKeys -> linked list of right indices.
    // Use a dedicated last slot for the sentinel key (0xFFFF_FFFF).
    const tableSize = nextPowerOfTwo(rightLen * 2);
    const mask = tableSize - 1;
    const tableLen = tableSize + 1;
    const tableKeys = new Uint32Array(tableLen);
    tableKeys.fill(EMPTY_KEY);
    const heads = new Uint32Array(tableLen);
    heads.fill(EMPTY_U32);
    const next = new Uint32Array(rightLen);

    for (let j = rightLen - 1; j >= 0; j--) {
      const key = rightU32[j];
      let slot;
      if (key === EMPTY_KEY) {
        slot = tableSize;
      } else {
        slot = (Math.imul(key, 2654435761) >>> 0) & mask;
        while (true) {
          const existing = tableKeys[slot];
          if (existing === EMPTY_KEY) {
            tableKeys[slot] = key;
            break;
          }
          if (existing === key) break;
          slot = (slot + 1) & mask;
        }
      }
      next[j] = heads[slot];
      heads[slot] = j;
    }

    /** @type {Uint32Array} */
    const leftHead = new Uint32Array(leftLen);
    leftHead.fill(EMPTY_U32);
    /** @type {Uint32Array} */
    const counts = new Uint32Array(leftLen);

    let total = 0;
    for (let i = 0; i < leftLen; i++) {
      const key = leftU32[i];
      let headPtr = EMPTY_U32;
      if (key === EMPTY_KEY) {
        headPtr = heads[tableSize];
      } else {
        let slot = (Math.imul(key, 2654435761) >>> 0) & mask;
        while (true) {
          const existing = tableKeys[slot];
          if (existing === key) {
            headPtr = heads[slot];
            break;
          }
          if (existing === EMPTY_KEY) break;
          slot = (slot + 1) & mask;
        }
      }

      leftHead[i] = headPtr;
      let c = 0;
      for (let ptr = headPtr; ptr !== EMPTY_U32; ptr = next[ptr]) c += 1;
      if (joinType === "left" && c === 0) c = 1;
      counts[i] = c;
      total += c;
    }

    if (total === 0) {
      return { leftIndex: new Uint32Array(), rightIndex: new Uint32Array() };
    }

    const leftIndex = new Uint32Array(total);
    const rightIndex = new Uint32Array(total);
    let p = 0;
    for (let i = 0; i < leftLen; i++) {
      const c = counts[i];
      if (c === 0) continue;
      const headPtr = leftHead[i];
      if (headPtr === EMPTY_U32) {
        leftIndex[p] = i;
        rightIndex[p] = EMPTY_U32;
        p += 1;
        continue;
      }
      for (let ptr = headPtr; ptr !== EMPTY_U32; ptr = next[ptr]) {
        leftIndex[p] = i;
        rightIndex[p] = ptr;
        p += 1;
      }
    }

    return { leftIndex, rightIndex };
  }
}
