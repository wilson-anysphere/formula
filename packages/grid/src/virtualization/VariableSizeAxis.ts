export interface AxisVisibleRange {
  start: number;
  end: number;
  offset: number;
}

function lowerBound(values: number[], target: number): number {
  let low = 0;
  let high = values.length;

  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (values[mid] < target) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }

  return low;
}

export class VariableSizeAxis {
  readonly defaultSize: number;

  // Monotonic counter that bumps whenever the effective sizing for any index changes.
  // Used by VirtualScrollManager to cheaply invalidate cached viewport computations.
  private version = 0;

  private overrides = new Map<number, number>();
  private overrideIndices: number[] = [];
  // Size diffs (override - default) aligned with `overrideIndices`.
  private diffs: number[] = [];
  // Fenwick tree over `diffs` for fast prefix sums (1-indexed).
  private diffBit: number[] = [];

  constructor(defaultSize: number) {
    if (!Number.isFinite(defaultSize) || defaultSize <= 0) {
      throw new Error(`defaultSize must be a positive finite number, got ${defaultSize}`);
    }
    this.defaultSize = defaultSize;
  }

  getVersion(): number {
    return this.version;
  }

  getSize(index: number): number {
    return this.overrides.get(index) ?? this.defaultSize;
  }

  /**
   * Replace the entire set of size overrides.
   *
   * This is intended for bulk updates (e.g. applying persisted row/col sizes) where calling
   * {@link setSize} repeatedly can devolve into O(n^2) behavior due to prefix-sum updates.
   *
   * The provided map may include entries equal to `defaultSize`; those are treated as "no override"
   * and are ignored.
   */
  setOverrides(overrides: ReadonlyMap<number, number>): void {
    const entries: Array<[number, number]> = [];
    for (const [index, size] of overrides) {
      if (!Number.isSafeInteger(index) || index < 0) {
        throw new Error(`index must be a non-negative safe integer, got ${index}`);
      }
      if (!Number.isFinite(size) || size <= 0) {
        throw new Error(`size must be a positive finite number, got ${size}`);
      }
      if (size === this.defaultSize) continue;
      entries.push([index, size]);
    }

    entries.sort((a, b) => a[0] - b[0]);

    // Fast path: if the normalized override set is identical to our current overrides, do nothing.
    if (entries.length === this.overrideIndices.length) {
      let identical = true;
      for (let i = 0; i < entries.length; i++) {
        const [index, size] = entries[i]!;
        if (this.overrideIndices[i] !== index || this.overrides.get(index) !== size) {
          identical = false;
          break;
        }
      }
      if (identical) return;
    }

    const nextOverrides = new Map<number, number>();
    const nextIndices = new Array<number>(entries.length);
    const nextDiffs = new Array<number>(entries.length);

    for (let i = 0; i < entries.length; i++) {
      const [index, size] = entries[i]!;
      nextOverrides.set(index, size);
      nextIndices[i] = index;
      nextDiffs[i] = size - this.defaultSize;
    }

    this.overrides = nextOverrides;
    this.overrideIndices = nextIndices;
    this.diffs = nextDiffs;
    this.rebuildDiffBit();
    this.version++;
  }

  setSize(index: number, size: number): void {
    if (!Number.isSafeInteger(index) || index < 0) {
      throw new Error(`index must be a non-negative safe integer, got ${index}`);
    }
    if (!Number.isFinite(size) || size <= 0) {
      throw new Error(`size must be a positive finite number, got ${size}`);
    }

    if (size === this.defaultSize) {
      this.deleteSize(index);
      return;
    }

    const existing = this.overrides.get(index);
    if (existing === size) return;
    this.overrides.set(index, size);

    const pos = lowerBound(this.overrideIndices, index);
    const diff = size - this.defaultSize;
    const previousDiff = existing === undefined ? 0 : existing - this.defaultSize;
    const delta = diff - previousDiff;

    if (this.overrideIndices[pos] !== index) {
      // New override inserted into the sorted list. This shifts later indices, so rebuild
      // the Fenwick tree from scratch (still O(n), but insertion itself already costs O(n)
      // due to array splices and is typically rare compared to "update existing override"
      // operations during resize drags).
      this.overrideIndices.splice(pos, 0, index);
      this.diffs.splice(pos, 0, diff);
      this.rebuildDiffBit();
      this.version++;
      return;
    }

    // Existing override updated in place; adjust the diff and update the Fenwick tree.
    this.diffs[pos] = diff;
    this.diffBitAdd(pos, delta);
    this.version++;
  }

  deleteSize(index: number): void {
    const existing = this.overrides.get(index);
    if (existing === undefined) return;
    this.overrides.delete(index);

    const pos = lowerBound(this.overrideIndices, index);
    if (this.overrideIndices[pos] !== index) return;

    this.overrideIndices.splice(pos, 1);
    this.diffs.splice(pos, 1);
    this.rebuildDiffBit();
    this.version++;
  }

  positionOf(index: number): number {
    if (!Number.isSafeInteger(index) || index < 0) {
      throw new Error(`index must be a non-negative safe integer, got ${index}`);
    }

    const diff = this.diffBefore(index);
    return index * this.defaultSize + diff;
  }

  totalSize(count: number): number {
    if (!Number.isSafeInteger(count) || count < 0) {
      throw new Error(`count must be a non-negative safe integer, got ${count}`);
    }
    return count * this.defaultSize + this.diffBefore(count);
  }

  indexAt(position: number, options?: { min?: number; maxInclusive?: number }): number {
    const min = options?.min ?? 0;
    if (!Number.isSafeInteger(min) || min < 0) {
      throw new Error(`min must be a non-negative safe integer, got ${min}`);
    }

    if (!Number.isFinite(position)) {
      throw new Error(`position must be finite, got ${position}`);
    }

    const minPos = this.positionOf(min);
    if (position <= minPos) return min;

    const maxInclusive = options?.maxInclusive;
    if (maxInclusive !== undefined) {
      if (!Number.isSafeInteger(maxInclusive) || maxInclusive < min) {
        throw new Error(
          `maxInclusive must be a safe integer >= min (${min}), got ${maxInclusive}`
        );
      }

      const totalPos = this.positionOf(maxInclusive + 1);
      if (position >= totalPos) return maxInclusive;

      let low = min;
      let high = maxInclusive;
      while (low < high) {
        const mid = Math.floor((low + high + 1) / 2);
        if (this.positionOf(mid) <= position) low = mid;
        else high = mid - 1;
      }
      return low;
    }

    let high = Math.max(min + 1, Math.floor(position / this.defaultSize) + 1);
    while (this.positionOf(high) <= position) {
      high = high * 2;
      if (!Number.isSafeInteger(high) || high > Number.MAX_SAFE_INTEGER / 2) {
        break;
      }
    }

    let low = min;
    let highInclusive = high - 1;
    while (low < highInclusive) {
      const mid = Math.floor((low + highInclusive + 1) / 2);
      if (this.positionOf(mid) <= position) low = mid;
      else highInclusive = mid - 1;
    }
    return low;
  }

  visibleRange(scroll: number, viewportSize: number, options?: { min?: number; maxExclusive?: number }): AxisVisibleRange {
    const min = options?.min ?? 0;
    const maxExclusive = options?.maxExclusive;
    const maxInclusive = maxExclusive === undefined ? undefined : Math.max(min, maxExclusive - 1);

    const start = this.indexAt(scroll, { min, maxInclusive });
    const startPos = this.positionOf(start);
    const offset = scroll - startPos;

    let end = start;
    let remaining = viewportSize + offset;

    const endLimit = maxExclusive ?? Number.MAX_SAFE_INTEGER;
    while (remaining > 0 && end < endLimit) {
      remaining -= this.getSize(end);
      end++;
    }

    return { start, end, offset };
  }

  private diffBefore(index: number): number {
    const pos = lowerBound(this.overrideIndices, index);
    if (pos === 0) return 0;
    return this.diffBitSum(pos);
  }

  private rebuildDiffBit(): void {
    const n = this.diffs.length;
    const bit = new Array<number>(n + 1).fill(0);

    // Build in O(n): write raw values then propagate partial sums to parent buckets.
    for (let i = 0; i < n; i++) {
      bit[i + 1] = this.diffs[i]!;
    }
    for (let i = 1; i <= n; i++) {
      const j = i + (i & -i);
      if (j <= n) bit[j] += bit[i]!;
    }

    this.diffBit = bit;
  }

  private diffBitAdd(pos0: number, delta: number): void {
    if (delta === 0) return;
    const bit = this.diffBit;
    for (let i = pos0 + 1; i < bit.length; i += i & -i) {
      bit[i] = (bit[i] ?? 0) + delta;
    }
  }

  /**
   * @param count Number of leading entries to sum (i.e. sum diffs[0..count-1]).
   */
  private diffBitSum(count: number): number {
    const bit = this.diffBit;
    let sum = 0;
    for (let i = count; i > 0; i -= i & -i) {
      sum += bit[i] ?? 0;
    }
    return sum;
  }
}
