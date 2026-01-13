/**
 * @param {number} seed
 * @returns {() => number}
 */
export function createSeededRng(seed) {
  // Mulberry32
  let a = seed >>> 0;
  return function rng() {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * Reservoir sampling indices without replacement.
 * @param {number} total
 * @param {number} sampleSize
 * @param {() => number} rng
 */
export function randomSampleIndices(total, sampleSize, rng) {
  if (!Number.isInteger(total) || total < 0) throw new Error(`total must be a non-negative integer, got: ${total}`);
  if (!Number.isInteger(sampleSize) || sampleSize < 0) {
    throw new Error(`sampleSize must be a non-negative integer, got: ${sampleSize}`);
  }

  if (sampleSize === 0 || total === 0) return [];
  if (sampleSize >= total) return Array.from({ length: total }, (_, i) => i);

  const reservoir = Array.from({ length: sampleSize }, (_, i) => i);
  for (let i = sampleSize; i < total; i++) {
    const j = Math.floor(rng() * (i + 1));
    if (j < sampleSize) reservoir[j] = i;
  }

  reservoir.sort((a, b) => a - b);
  return reservoir;
}

/**
 * @template T
 * @param {T[]} rows
 * @param {number} sampleSize
 * @param {{ seed?: number, rng?: () => number }} [options]
 */
export function randomSampleRows(rows, sampleSize, options = {}) {
  const rng = options.rng ?? createSeededRng(options.seed ?? 1);
  const indices = randomSampleIndices(rows.length, sampleSize, rng);
  return indices.map((i) => rows[i]);
}

/**
 * @template T
 * @param {T[]} rows
 * @param {number} sampleSize
 */
export function headSampleRows(rows, sampleSize) {
  if (sampleSize <= 0 || rows.length === 0) return [];
  if (sampleSize >= rows.length) return rows.slice();
  return rows.slice(0, sampleSize);
}

/**
 * @template T
 * @param {T[]} rows
 * @param {number} sampleSize
 */
export function tailSampleRows(rows, sampleSize) {
  if (sampleSize <= 0 || rows.length === 0) return [];
  if (sampleSize >= rows.length) return rows.slice();
  return rows.slice(rows.length - sampleSize);
}

/**
 * Systematic (evenly spaced) sampling without replacement.
 *
 * The starting offset is derived deterministically from the provided `seed` / `rng`
 * unless `options.offset` is provided.
 *
 * @template T
 * @param {T[]} rows
 * @param {number} sampleSize
 * @param {{ seed?: number, rng?: () => number, offset?: number }} [options]
 */
export function systematicSampleRows(rows, sampleSize, options = {}) {
  if (sampleSize <= 0 || rows.length === 0) return [];
  if (sampleSize >= rows.length) return rows.slice();

  const rng = options.rng ?? createSeededRng(options.seed ?? 1);
  const offsetRaw = options.offset ?? rng();
  // Force the offset into [0, 1) so callers can pass e.g. 1.25 to mean 0.25.
  const offset = ((offsetRaw % 1) + 1) % 1;

  const step = rows.length / sampleSize;
  /** @type {T[]} */
  const out = [];
  for (let i = 0; i < sampleSize; i++) {
    // `step >= 1` because `sampleSize < rows.length` above. This guarantees indices are unique.
    let idx = Math.floor((i + offset) * step);
    if (idx >= rows.length) idx = rows.length - 1;
    out.push(rows[idx]);
  }
  return out;
}

/**
 * @typedef {{ score: number, key: string }} HeapItem
 */

/**
 * @param {HeapItem[]} heap
 * @param {number} i
 * @param {number} j
 */
function heapSwap(heap, i, j) {
  const tmp = heap[i];
  heap[i] = heap[j];
  heap[j] = tmp;
}

/**
 * @param {HeapItem[]} heap
 * @param {number} idx
 */
function heapSiftUp(heap, idx) {
  while (idx > 0) {
    const parent = (idx - 1) >> 1;
    if (heap[parent].score <= heap[idx].score) break;
    heapSwap(heap, parent, idx);
    idx = parent;
  }
}

/**
 * @param {HeapItem[]} heap
 * @param {number} idx
 */
function heapSiftDown(heap, idx) {
  for (;;) {
    const left = idx * 2 + 1;
    const right = idx * 2 + 2;
    let smallest = idx;
    if (left < heap.length && heap[left].score < heap[smallest].score) smallest = left;
    if (right < heap.length && heap[right].score < heap[smallest].score) smallest = right;
    if (smallest === idx) break;
    heapSwap(heap, idx, smallest);
    idx = smallest;
  }
}

/**
 * @param {HeapItem[]} heap
 * @param {HeapItem} item
 */
function heapPush(heap, item) {
  heap.push(item);
  heapSiftUp(heap, heap.length - 1);
}

/**
 * @param {HeapItem[]} heap
 * @param {HeapItem} item
 */
function heapReplaceMin(heap, item) {
  heap[0] = item;
  heapSiftDown(heap, 0);
}

/**
 * Select `sampleSize` unique stratum keys with probability proportional to stratum size
 * (weighted sampling without replacement) without constructing an O(totalRows) helper array.
 *
 * Implementation: Efraimidis–Spirakis "A-Res" weighted reservoir sampling.
 *
 * @param {Iterable<[string, number]>} stratumEntries
 * @param {number} sampleSize
 * @param {() => number} rng
 * @returns {string[]}
 */
function sampleStrataWithoutReplacement(stratumEntries, sampleSize, rng) {
  if (sampleSize <= 0) return [];

  /** @type {HeapItem[]} */
  const heap = [];

  for (const [key, weight] of stratumEntries) {
    if (weight <= 0) continue;
    // rng() is in [0, 1). Convert to (0, 1] to avoid generating `Infinity` in log-based variants.
    const u = 1 - rng();
    const score = Math.pow(u, 1 / weight);

    if (heap.length < sampleSize) {
      heapPush(heap, { score, key });
      continue;
    }
    if (score > heap[0].score) {
      heapReplaceMin(heap, { score, key });
    }
  }

  return heap.map((item) => item.key);
}

/**
 * @template T
 * @param {T[]} rows
 * @param {number} sampleSize
 * @param {{ getStratum: (row: T) => string, seed?: number, rng?: () => number }} options
 */
export function stratifiedSampleRows(rows, sampleSize, options) {
  const { getStratum } = options;
  const rng = options.rng ?? createSeededRng(options.seed ?? 1);

  if (sampleSize <= 0) return [];
  if (rows.length === 0) return [];
  if (sampleSize >= rows.length) return rows.slice();

  /** @type {Map<string, number>} */
  const strata = new Map();
  for (let i = 0; i < rows.length; i++) {
    const key = getStratum(rows[i]);
    strata.set(key, (strata.get(key) ?? 0) + 1);
  }

  const strataCount = strata.size;

  /** @type {Map<string, number>} */
  const allocation = new Map();

  if (sampleSize < strataCount) {
    // Not enough room for every stratum: choose strata (weighted by size) and take one each.
    const chosen = sampleStrataWithoutReplacement(strata.entries(), sampleSize, rng);
    for (const key of chosen) allocation.set(key, 1);
  } else {
    // Ensure at least one sample per stratum, then distribute the remainder proportionally.
    let remaining = sampleSize - strataCount;

    const totalRemainderPopulation = rows.length - strataCount;
    if (remaining > 0 && totalRemainderPopulation > 0) {
      const shares = [];
      for (const [key, count] of strata.entries()) {
        allocation.set(key, 1);
        const available = Math.max(count - 1, 0);
        const exact = (available / totalRemainderPopulation) * remaining;
        shares.push({
          key,
          count,
          available,
          exact,
          floor: Math.floor(exact),
          frac: exact - Math.floor(exact),
        });
      }

      let allocated = 0;
      for (const share of shares) {
        const add = Math.min(share.available, share.floor);
        allocation.set(share.key, (allocation.get(share.key) ?? 0) + add);
        allocated += add;
      }

      remaining -= allocated;
      shares.sort((a, b) => b.frac - a.frac);
      for (const share of shares) {
        if (remaining <= 0) break;
        const current = allocation.get(share.key) ?? 0;
        if (current >= share.count) continue;
        allocation.set(share.key, current + 1);
        remaining--;
      }

      // Any remaining samples are distributed randomly to strata that still have capacity.
      while (remaining > 0) {
        const candidates = shares.filter((s) => (allocation.get(s.key) ?? 0) < s.count);
        if (candidates.length === 0) break;
        const pick = candidates[Math.floor(rng() * candidates.length)];
        allocation.set(pick.key, (allocation.get(pick.key) ?? 0) + 1);
        remaining--;
      }
    } else {
      // `remaining <= 0` or no remainder population: still ensure we allocate one sample per stratum.
      for (const [key] of strata.entries()) allocation.set(key, 1);
    }
  }

  /** @type {Map<string, { k: number, seen: number, reservoir: number[] }>} */
  const reservoirs = new Map();
  for (const [key, k] of allocation.entries()) {
    if (k > 0) reservoirs.set(key, { k, seen: 0, reservoir: [] });
  }

  // Second pass: per-stratum reservoir sampling of global row indices.
  for (let i = 0; i < rows.length; i++) {
    const key = getStratum(rows[i]);
    const state = reservoirs.get(key);
    if (!state) continue;

    const seen = state.seen;
    if (seen < state.k) {
      state.reservoir.push(i);
      state.seen = seen + 1;
      continue;
    }

    const j = Math.floor(rng() * (seen + 1));
    if (j < state.k) state.reservoir[j] = i;
    state.seen = seen + 1;
  }

  /** @type {number[]} */
  const sampledIndices = [];
  for (const state of reservoirs.values()) {
    for (const idx of state.reservoir) sampledIndices.push(idx);
  }

  sampledIndices.sort((a, b) => a - b);
  return sampledIndices.map((i) => rows[i]);
}
