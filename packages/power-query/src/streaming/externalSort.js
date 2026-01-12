import { PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../values.js";

/**
 * @typedef {import("./spillStore.js").SpillStore} SpillStore
 */

/**
 * @param {AbortSignal | undefined} signal
 */
function throwIfAborted(signal) {
  if (!signal?.aborted) return;
  const err = new Error("Aborted");
  err.name = "AbortError";
  throw err;
}

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {unknown} value
 * @returns {boolean}
 */
function isNullish(value) {
  return value == null;
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function normalizeMissing(value) {
  return value === undefined ? null : value;
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function valueToString(value) {
  if (value == null) return "";
  if (isDate(value)) return value.toISOString();
  return String(value);
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @returns {number}
 */
export function compareNonNull(a, b) {
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "boolean" && typeof b === "boolean") return Number(a) - Number(b);
  if (isDate(a) && isDate(b)) return a.getTime() - b.getTime();
  if (a instanceof PqDateTimeZone && b instanceof PqDateTimeZone) return a.toDate().getTime() - b.toDate().getTime();
  if (a instanceof PqTime && b instanceof PqTime) return a.milliseconds - b.milliseconds;
  if (a instanceof PqDuration && b instanceof PqDuration) return a.milliseconds - b.milliseconds;
  if (a instanceof PqDecimal && b instanceof PqDecimal) {
    const aNum = Number(a.value);
    const bNum = Number(b.value);
    if (Number.isFinite(aNum) && Number.isFinite(bNum)) return aNum - bNum;
  }
  return valueToString(a).localeCompare(valueToString(b));
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @param {{ direction: "ascending" | "descending", nulls: "first" | "last" }} options
 * @returns {number}
 */
export function compareValues(a, b, options) {
  const aNull = isNullish(a);
  const bNull = isNullish(b);
  if (aNull && bNull) return 0;
  if (aNull || bNull) {
    const nullCmp = aNull ? -1 : 1;
    const adjusted = options.nulls === "first" ? nullCmp : -nullCmp;
    return adjusted;
  }

  const baseCmp = compareNonNull(a, b);
  return options.direction === "ascending" ? baseCmp : -baseCmp;
}

class MinHeap {
  /**
   * @param {(a: any, b: any) => number} compare
   */
  constructor(compare) {
    this.compare = compare;
    /** @type {any[]} */
    this.items = [];
  }

  get size() {
    return this.items.length;
  }

  /**
   * @param {any} item
   */
  push(item) {
    const items = this.items;
    items.push(item);
    let idx = items.length - 1;
    while (idx > 0) {
      const parent = Math.floor((idx - 1) / 2);
      if (this.compare(items[idx], items[parent]) >= 0) break;
      const tmp = items[idx];
      items[idx] = items[parent];
      items[parent] = tmp;
      idx = parent;
    }
  }

  pop() {
    const items = this.items;
    if (items.length === 0) return null;
    const out = items[0];
    const last = items.pop();
    if (items.length > 0) {
      items[0] = last;
      let idx = 0;
      while (true) {
        const left = idx * 2 + 1;
        const right = left + 1;
        let smallest = idx;
        if (left < items.length && this.compare(items[left], items[smallest]) < 0) smallest = left;
        if (right < items.length && this.compare(items[right], items[smallest]) < 0) smallest = right;
        if (smallest === idx) break;
        const tmp = items[idx];
        items[idx] = items[smallest];
        items[smallest] = tmp;
        idx = smallest;
      }
    }
    return out;
  }
}

/**
 * @param {AsyncIterable<unknown[][]>} batches
 * @returns {AsyncIterable<unknown[]>}
 */
export async function* rowsFromBatches(batches) {
  for await (const batch of batches) {
    if (!Array.isArray(batch) || batch.length === 0) continue;
    for (const row of batch) {
      if (Array.isArray(row)) yield row;
    }
  }
}

/**
 * External sort an input stream with bounded memory.
 *
 * The implementation writes sorted runs into `store` when thresholds are exceeded, then performs
 * a k-way merge of those runs.
 *
 * @param {AsyncIterable<unknown[][]>} inputBatches
 * @param {(a: unknown[], b: unknown[]) => number} comparator
 * @param {{
 *   store: SpillStore;
 *   runKeyPrefix: string;
 *   batchSize: number;
 *   maxInMemoryRows: number;
 *   maxInMemoryBytes?: number;
 *   signal?: AbortSignal;
 *   onSpill?: (details: { runCount: number }) => void;
 * }} options
 * @returns {AsyncIterable<unknown[][]>}
 */
export async function* externalSortBatches(inputBatches, comparator, options) {
  const store = options.store;
  const runKeyPrefix = options.runKeyPrefix;
  const batchSize = options.batchSize;
  const maxInMemoryRows = Math.max(1, Math.trunc(options.maxInMemoryRows));
  const maxInMemoryBytes =
    typeof options.maxInMemoryBytes === "number" && Number.isFinite(options.maxInMemoryBytes) && options.maxInMemoryBytes > 0
      ? options.maxInMemoryBytes
      : null;
  const signal = options.signal;

  /** @type {unknown[][]} */
  let buffer = [];
  let bufferBytes = 0;
  /** @type {string[]} */
  const runKeys = [];

  const flushRun = async () => {
    if (buffer.length === 0) return;
    buffer.sort(comparator);
    const runId = runKeys.length;
    const runKey = `${runKeyPrefix}:run:${runId}`;
    for (let i = 0; i < buffer.length; i += batchSize) {
      await store.putBatch(runKey, buffer.slice(i, i + batchSize));
    }
    runKeys.push(runKey);
    buffer = [];
    bufferBytes = 0;
    options.onSpill?.({ runCount: runKeys.length });
  };

  try {
    for await (const inBatch of inputBatches) {
      throwIfAborted(signal);
      if (!Array.isArray(inBatch) || inBatch.length === 0) continue;
      for (const row of inBatch) {
        if (!Array.isArray(row)) continue;
        buffer.push(row.map(normalizeMissing));
        if (maxInMemoryBytes != null) {
          try {
            bufferBytes += JSON.stringify(row, (_k, v) => (typeof v === "bigint" ? v.toString() : v)).length;
          } catch {
            bufferBytes += 16;
          }
        }
        if (buffer.length >= maxInMemoryRows || (maxInMemoryBytes != null && bufferBytes >= maxInMemoryBytes)) {
          await flushRun();
        }
      }
    }

    // In-memory fast path when nothing spilled.
    if (runKeys.length === 0) {
      if (buffer.length === 0) return;
      buffer.sort(comparator);
      for (let i = 0; i < buffer.length; i += batchSize) {
        yield buffer.slice(i, i + batchSize);
      }
      return;
    }

    await flushRun();

    /** @type {Array<AsyncIterator<unknown[]>>} */
    const iterators = [];
    for (const key of runKeys) {
      iterators.push(store.iterateRows(key)[Symbol.asyncIterator]());
    }

    const heap = new MinHeap((a, b) => {
      const cmp = comparator(a.row, b.row);
      if (cmp !== 0) return cmp;
      return a.run - b.run;
    });

    for (let run = 0; run < iterators.length; run++) {
      throwIfAborted(signal);
      const next = await iterators[run].next();
      if (next.done) continue;
      heap.push({ run, row: /** @type {unknown[]} */ (next.value) });
    }

    /** @type {unknown[][]} */
    let outBatch = [];

    while (heap.size > 0) {
      throwIfAborted(signal);
      const item = heap.pop();
      if (!item) break;
      outBatch.push(item.row);
      if (outBatch.length >= batchSize) {
        yield outBatch;
        outBatch = [];
      }

      const iterator = iterators[item.run];
      const next = await iterator.next();
      if (!next.done) {
        heap.push({ run: item.run, row: /** @type {unknown[]} */ (next.value) });
      }
    }

    if (outBatch.length > 0) yield outBatch;
  } finally {
    await Promise.all(
      runKeys.map(async (key) => {
        try {
          await store.clear(key);
        } catch {
          // ignore
        }
      }),
    );
    try {
      await store.clearPrefix(runKeyPrefix);
    } catch {
      // ignore
    }
  }
}

