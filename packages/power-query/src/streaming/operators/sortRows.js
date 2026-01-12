import { externalSortBatches, compareValues } from "../externalSort.js";
import { makeSpillKeyPrefix } from "../spillStore.js";

/**
 * @typedef {import("../../model.js").SortSpec} SortSpec
 * @typedef {import("../spillStore.js").SpillStore} SpillStore
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
 * @param {import("../../table.js").Column[]} columns
 */
function buildColumnIndex(columns) {
  return new Map(columns.map((c, idx) => [c.name, idx]));
}

/**
 * Streaming external sort for `sortRows`.
 *
 * @param {{
 *   columns: import("../../table.js").Column[];
 *   batches: AsyncIterable<unknown[][]>;
 * }} input
 * @param {SortSpec[]} sortBy
 * @param {{
 *   store: SpillStore;
 *   batchSize: number;
 *   maxInMemoryRows: number;
 *   maxInMemoryBytes?: number;
 *   signal?: AbortSignal;
 *   onProgress?: (event: any) => void;
 *   queryId?: string;
 * }} options
 */
export function sortRowsStreaming(input, sortBy, options) {
  if (!Array.isArray(sortBy) || sortBy.length === 0) {
    return { columns: input.columns, batches: input.batches };
  }

  const columnIndex = buildColumnIndex(input.columns);
  const getIndex = (name) => {
    const idx = columnIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${input.columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  };
  const specs = sortBy.map((spec) => ({
    idx: getIndex(spec.column),
    direction: spec.direction ?? "ascending",
    nulls: spec.nulls ?? "last",
  }));

  const runKeyPrefix = makeSpillKeyPrefix("pq-sort");
  let seq = 0;

  /** @type {AsyncIterable<unknown[][]>} */
  const decorated = (async function* () {
    for await (const batch of input.batches) {
      throwIfAborted(options.signal);
      if (!Array.isArray(batch) || batch.length === 0) continue;
      const out = new Array(batch.length);
      for (let i = 0; i < batch.length; i++) {
        const row = Array.isArray(batch[i]) ? batch[i] : [];
        out[i] = [...row, seq++];
      }
      yield out;
    }
  })();

  const comparator = (a, b) => {
    for (const spec of specs) {
      const idx = spec.idx;
      const cmp = compareValues(a[idx], b[idx], spec);
      if (cmp !== 0) return cmp;
    }
    // Stability: preserve original order.
    return (a[a.length - 1] ?? 0) - (b[b.length - 1] ?? 0);
  };

  let didSpill = false;
  const sortedDecorated = externalSortBatches(decorated, comparator, {
    store: options.store,
    runKeyPrefix,
    batchSize: options.batchSize,
    maxInMemoryRows: options.maxInMemoryRows,
    maxInMemoryBytes: options.maxInMemoryBytes,
    signal: options.signal,
    onSpill: ({ runCount }) => {
      if (runCount > 0) didSpill = true;
      options.onProgress?.({
        type: "stream:spill",
        queryId: options.queryId ?? "<unknown>",
        operator: "sortRows",
        runCount,
      });
    },
  });

  /** @type {AsyncIterable<unknown[][]>} */
  const batches = (async function* () {
    for await (const batch of sortedDecorated) {
      throwIfAborted(options.signal);
      if (!Array.isArray(batch) || batch.length === 0) continue;
      yield batch.map((row) => row.slice(0, row.length - 1));
    }
    if (didSpill) {
      options.onProgress?.({
        type: "stream:operator",
        queryId: options.queryId ?? "<unknown>",
        operator: "sortRows",
        spilled: true,
      });
    }
  })();

  return { columns: input.columns, batches };
}
