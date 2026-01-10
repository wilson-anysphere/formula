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

  private overrides = new Map<number, number>();
  private overrideIndices: number[] = [];
  private prefixDiffs: number[] = [];

  constructor(defaultSize: number) {
    if (!Number.isFinite(defaultSize) || defaultSize <= 0) {
      throw new Error(`defaultSize must be a positive finite number, got ${defaultSize}`);
    }
    this.defaultSize = defaultSize;
  }

  getSize(index: number): number {
    return this.overrides.get(index) ?? this.defaultSize;
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
    this.overrides.set(index, size);

    const pos = lowerBound(this.overrideIndices, index);
    if (this.overrideIndices[pos] !== index) {
      this.overrideIndices.splice(pos, 0, index);
      this.prefixDiffs.splice(pos, 0, 0);
    }

    const diff = size - this.defaultSize;
    const previousDiff = existing === undefined ? 0 : existing - this.defaultSize;
    const delta = diff - previousDiff;

    for (let i = pos; i < this.prefixDiffs.length; i++) {
      this.prefixDiffs[i] += delta;
    }
  }

  deleteSize(index: number): void {
    if (!this.overrides.has(index)) return;
    const existing = this.overrides.get(index);
    this.overrides.delete(index);
    if (existing === undefined) return;

    const pos = lowerBound(this.overrideIndices, index);
    if (this.overrideIndices[pos] !== index) return;

    this.overrideIndices.splice(pos, 1);
    const diff = existing - this.defaultSize;
    this.prefixDiffs.splice(pos, 1);
    for (let i = pos; i < this.prefixDiffs.length; i++) {
      this.prefixDiffs[i] -= diff;
    }
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
    return this.prefixDiffs[pos - 1];
  }
}

