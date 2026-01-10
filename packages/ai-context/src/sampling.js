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
 * @param {{ getStratum: (row: T) => string, seed?: number, rng?: () => number }} options
 */
export function stratifiedSampleRows(rows, sampleSize, options) {
  const { getStratum } = options;
  const rng = options.rng ?? createSeededRng(options.seed ?? 1);

  if (sampleSize <= 0) return [];
  if (rows.length === 0) return [];
  if (sampleSize >= rows.length) return rows.slice();

  /** @type {Map<string, number[]>} */
  const strata = new Map();
  for (let i = 0; i < rows.length; i++) {
    const key = getStratum(rows[i]);
    const bucket = strata.get(key);
    if (bucket) bucket.push(i);
    else strata.set(key, [i]);
  }

  const stratumEntries = [...strata.entries()];

  /** @type {Map<string, number>} */
  const allocation = new Map();

  if (sampleSize < stratumEntries.length) {
    // Not enough room for every stratum: choose strata (weighted by size) and take one each.
    const weighted = stratumEntries.flatMap(([key, indices]) => Array.from({ length: indices.length }, () => key));
    const chosen = new Set();
    while (chosen.size < sampleSize) {
      const key = weighted[Math.floor(rng() * weighted.length)];
      chosen.add(key);
    }
    for (const key of chosen) allocation.set(key, 1);
  } else {
    // Ensure at least one sample per stratum, then distribute the remainder proportionally.
    for (const [key] of stratumEntries) allocation.set(key, 1);
    let remaining = sampleSize - stratumEntries.length;

    const totalRemainderPopulation = rows.length - stratumEntries.length;
    if (remaining > 0 && totalRemainderPopulation > 0) {
      const shares = stratumEntries.map(([key, indices]) => {
        const available = Math.max(indices.length - 1, 0);
        const exact = (available / totalRemainderPopulation) * remaining;
        return { key, available, exact, floor: Math.floor(exact), frac: exact - Math.floor(exact) };
      });

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
        if (current >= share.available + 1) continue;
        allocation.set(share.key, current + 1);
        remaining--;
      }

      // Any remaining samples are distributed randomly to strata that still have capacity.
      while (remaining > 0) {
        const candidates = shares.filter((s) => (allocation.get(s.key) ?? 0) < s.available + 1);
        if (candidates.length === 0) break;
        const pick = candidates[Math.floor(rng() * candidates.length)];
        allocation.set(pick.key, (allocation.get(pick.key) ?? 0) + 1);
        remaining--;
      }
    }
  }

  /** @type {number[]} */
  const sampledIndices = [];
  for (const [key, indices] of stratumEntries) {
    const n = allocation.get(key) ?? 0;
    if (n <= 0) continue;
    const chosen = randomSampleIndices(indices.length, Math.min(n, indices.length), rng);
    for (const pos of chosen) sampledIndices.push(indices[pos]);
  }

  sampledIndices.sort((a, b) => a - b);
  return sampledIndices.map((i) => rows[i]);
}
