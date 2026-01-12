import { valueKey } from "../../valueKey.js";
import { PqDecimal } from "../../values.js";
import { compareNonNull, externalSortBatches } from "../externalSort.js";
import { makeSpillKeyPrefix } from "../spillStore.js";

/**
 * @typedef {import("../../model.js").Aggregation} Aggregation
 * @typedef {import("../spillStore.js").SpillStore} SpillStore
 */

const NUMBER_TEXT_RE = /^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$/;

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function normalizeMissing(value) {
  return value === undefined ? null : value;
}

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
 * @returns {number | null}
 */
function toNumberOrNull(value) {
  if (value instanceof PqDecimal) {
    const parsed = Number(value.value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    if (!NUMBER_TEXT_RE.test(trimmed)) return null;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

/**
 * @param {import("../../table.js").Column[]} columns
 */
function buildColumnIndex(columns) {
  return new Map(columns.map((c, idx) => [c.name, idx]));
}

/**
 * Streaming groupBy using external sort + aggregation.
 *
 * This preserves materialized group ordering by sorting the final group results by each group's first-seen row index.
 *
 * @param {{
 *   columns: import("../../table.js").Column[];
 *   batches: AsyncIterable<unknown[][]>;
 * }} input
 * @param {string[]} groupColumns
 * @param {Aggregation[]} aggregations
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
export function groupByStreaming(input, groupColumns, aggregations, options) {
  const columnIndex = buildColumnIndex(input.columns);

  const getIndex = (name) => {
    const idx = columnIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${input.columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  };

  const groupIdx = groupColumns.map(getIndex);
  const aggSpecs = aggregations.map((agg) => ({
    ...agg,
    idx: getIndex(agg.column),
    as: agg.as ?? `${agg.op} of ${agg.column}`,
  }));

  for (const agg of aggSpecs) {
    if (!["count", "sum", "min", "max", "average", "countDistinct"].includes(agg.op)) {
      throw new Error(`Unsupported aggregation '${agg.op}'`);
    }
  }

  const resultColumns = [
    ...groupIdx.map((idx) => ({ ...input.columns[idx], type: input.columns[idx]?.type ?? "any" })),
    ...aggSpecs.map((agg) => ({
      name: agg.as,
      type:
        agg.op === "sum" || agg.op === "average" || agg.op === "count" || agg.op === "countDistinct"
          ? "number"
          : input.columns[agg.idx]?.type ?? "any",
    })),
  ];

  const recordRunKeyPrefix = makeSpillKeyPrefix("pq-groupby-records");
  const groupedRunKeyPrefix = makeSpillKeyPrefix("pq-groupby-groups");

  /** @type {AsyncIterable<unknown[][]>} */
  const records = (async function* () {
    let rowIndex = 0;
    for await (const batch of input.batches) {
      throwIfAborted(options.signal);
      if (!Array.isArray(batch) || batch.length === 0) continue;
      const out = new Array(batch.length);
      for (let i = 0; i < batch.length; i++) {
        const row = Array.isArray(batch[i]) ? batch[i] : [];
        const keyValues = groupIdx.map((idx) => row[idx]);
        const key = JSON.stringify(keyValues.map((v) => valueKey(normalizeMissing(v))));
        const aggValues = aggSpecs.map((agg) => row[agg.idx]);
        out[i] = [key, rowIndex++, ...keyValues, ...aggValues];
      }
      yield out;
    }
  })();

  const sortedRecords = externalSortBatches(records, (a, b) => {
    const keyA = String(a[0] ?? "");
    const keyB = String(b[0] ?? "");
    const keyCmp = keyA.localeCompare(keyB);
    if (keyCmp !== 0) return keyCmp;
    return (a[1] ?? 0) - (b[1] ?? 0);
  }, {
    store: options.store,
    runKeyPrefix: recordRunKeyPrefix,
    batchSize: options.batchSize,
    maxInMemoryRows: options.maxInMemoryRows,
    maxInMemoryBytes: options.maxInMemoryBytes,
    signal: options.signal,
    onSpill: ({ runCount }) => {
      options.onProgress?.({
        type: "stream:spill",
        queryId: options.queryId ?? "<unknown>",
        operator: "groupBy",
        phase: "inputSort",
        runCount,
      });
    },
  });

  /** @type {AsyncIterable<unknown[][]>} */
  const groupedRows = (async function* () {
    /** @type {string | null} */
    let currentKey = null;
    /** @type {number} */
    let firstRowIndex = 0;
    /** @type {unknown[]} */
    let keyValues = [];
    /** @type {any[]} */
    let states = [];

    /**
     * @returns {any[]}
     */
    const initStates = () =>
      aggSpecs.map((agg) => {
        switch (agg.op) {
          case "sum":
          case "average":
            return { sum: 0, count: 0 };
          case "count":
            return { count: 0 };
          case "min":
          case "max":
            return { value: null, has: false };
          case "countDistinct":
            return { set: new Set() };
          default:
            return { count: 0 };
        }
      });

    /**
     * @returns {unknown[] | null}
     */
    const finalizeGroup = () => {
      if (currentKey == null) return null;
      const row = [firstRowIndex, ...keyValues];
      states.forEach((state, idx) => {
        const agg = aggSpecs[idx];
        switch (agg.op) {
          case "sum":
            row.push(state.sum);
            break;
          case "average":
            row.push(state.count === 0 ? null : state.sum / state.count);
            break;
          case "count":
            row.push(state.count);
            break;
          case "min":
          case "max":
            row.push(state.has ? state.value : null);
            break;
          case "countDistinct":
            row.push(state.set.size);
            break;
          default:
            row.push(null);
            break;
        }
      });
      return row;
    };

    /** @type {unknown[][]} */
    let outBatch = [];

    for await (const batch of sortedRecords) {
      throwIfAborted(options.signal);
      if (!Array.isArray(batch) || batch.length === 0) continue;
      for (const record of batch) {
        const key = String(record[0] ?? "");
        const recordRowIndex = /** @type {number} */ (record[1] ?? 0);
        const recordKeyValues = record.slice(2, 2 + groupIdx.length);
        const recordAggValues = record.slice(2 + groupIdx.length);

        if (currentKey == null) {
          currentKey = key;
          firstRowIndex = recordRowIndex;
          keyValues = recordKeyValues;
          states = initStates();
        } else if (key !== currentKey) {
          const doneRow = finalizeGroup();
          if (doneRow) outBatch.push(doneRow);
          if (outBatch.length >= options.batchSize) {
            yield outBatch;
            outBatch = [];
          }

          currentKey = key;
          firstRowIndex = recordRowIndex;
          keyValues = recordKeyValues;
          states = initStates();
        }

        states.forEach((state, idx) => {
          const agg = aggSpecs[idx];
          const value = normalizeMissing(recordAggValues[idx]);

          switch (agg.op) {
            case "sum": {
              const num = toNumberOrNull(value);
              if (num != null) state.sum += num;
              break;
            }
            case "average": {
              const num = toNumberOrNull(value);
              if (num != null) {
                state.sum += num;
                state.count += 1;
              }
              break;
            }
            case "count":
              state.count += 1;
              break;
            case "min":
              if (value == null) break;
              if (!state.has || compareNonNull(value, state.value) < 0) {
                state.value = value;
                state.has = true;
              }
              break;
            case "max":
              if (value == null) break;
              if (!state.has || compareNonNull(value, state.value) > 0) {
                state.value = value;
                state.has = true;
              }
              break;
            case "countDistinct":
              state.set.add(valueKey(value));
              break;
            default:
              break;
          }
        });
      }
    }

    const finalRow = finalizeGroup();
    if (finalRow) outBatch.push(finalRow);
    if (outBatch.length > 0) yield outBatch;
  })();

  const sortedGroups = externalSortBatches(groupedRows, (a, b) => (a[0] ?? 0) - (b[0] ?? 0), {
    store: options.store,
    runKeyPrefix: groupedRunKeyPrefix,
    batchSize: options.batchSize,
    maxInMemoryRows: options.maxInMemoryRows,
    maxInMemoryBytes: options.maxInMemoryBytes,
    signal: options.signal,
    onSpill: ({ runCount }) => {
      options.onProgress?.({
        type: "stream:spill",
        queryId: options.queryId ?? "<unknown>",
        operator: "groupBy",
        phase: "outputSort",
        runCount,
      });
    },
  });

  /** @type {AsyncIterable<unknown[][]>} */
  const batches = (async function* () {
    for await (const batch of sortedGroups) {
      throwIfAborted(options.signal);
      if (!Array.isArray(batch) || batch.length === 0) continue;
      yield batch.map((row) => row.slice(1));
    }
    options.onProgress?.({
      type: "stream:operator",
      queryId: options.queryId ?? "<unknown>",
      operator: "groupBy",
      spilled: options.store.stats.batchesWritten > 0,
    });
  })();

  return { columns: resultColumns, batches };
}

