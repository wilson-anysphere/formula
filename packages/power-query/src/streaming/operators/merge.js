import { DataTable, makeUniqueColumnNames } from "../../table.js";
import { valueKey } from "../../valueKey.js";
import { makeSpillKeyPrefix } from "../spillStore.js";

/**
 * @typedef {import("../../model.js").MergeOp} MergeOp
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
 * @param {unknown} comparer
 * @returns {boolean}
 */
function isIgnoreCaseComparer(comparer) {
  if (!comparer || typeof comparer !== "object" || Array.isArray(comparer)) return false;
  // @ts-ignore - runtime inspection
  const comparerName = typeof comparer.comparer === "string" ? comparer.comparer.toLowerCase() : "";
  // @ts-ignore - runtime inspection
  return comparer.caseSensitive === false || comparerName === "ordinalignorecase";
}

/**
 * @param {unknown[]} row
 * @param {number[]} keyIndices
 * @param {boolean[] | null} keyIgnoreCase
 */
function compositeKeyForRow(row, keyIndices, keyIgnoreCase) {
  const parts = keyIndices.map((idx, keyPos) => {
    let value = row[idx];
    if (keyIgnoreCase && keyIgnoreCase[keyPos] && typeof value === "string") {
      value = value.toLowerCase();
    }
    return valueKey(value);
  });
  return JSON.stringify(parts);
}

/**
 * Streaming merge/join operator (inner + left, flat + nested).
 *
 * @param {{
 *   left: { columns: import("../../table.js").Column[]; batches: AsyncIterable<unknown[][]> };
 *   right: { columns: import("../../table.js").Column[]; batches: AsyncIterable<unknown[][]> };
 *   op: MergeOp;
 *   store: SpillStore;
 *   batchSize: number;
 *   maxInMemoryRows: number;
 *   maxInMemoryBytes?: number;
 *   signal?: AbortSignal;
 *   onProgress?: (event: any) => void;
 *   queryId?: string;
 * }} options
 */
export function mergeStreaming(options) {
  const left = options.left;
  const right = options.right;
  const op = options.op;

  const joinType = op.joinType ?? "inner";
  if (joinType !== "inner" && joinType !== "left") {
    throw new Error(`Streaming merge only supports joinType 'inner' and 'left' (got '${joinType}')`);
  }

  const joinMode = op.joinMode ?? "flat";

  const leftKeys =
    Array.isArray(op.leftKeys) && op.leftKeys.length > 0
      ? op.leftKeys
      : typeof op.leftKey === "string" && op.leftKey
        ? [op.leftKey]
        : [];
  const rightKeys =
    Array.isArray(op.rightKeys) && op.rightKeys.length > 0
      ? op.rightKeys
      : typeof op.rightKey === "string" && op.rightKey
        ? [op.rightKey]
        : [];

  if (leftKeys.length === 0 || rightKeys.length === 0) {
    throw new Error("merge requires join key columns");
  }
  if (leftKeys.length !== rightKeys.length) {
    throw new Error(`merge requires leftKeys/rightKeys to have the same length (got ${leftKeys.length} and ${rightKeys.length})`);
  }

  /** @type {boolean[] | null} */
  let keyIgnoreCase = null;
  if (Array.isArray(op.comparers) && op.comparers.length > 0) {
    if (op.comparers.length !== leftKeys.length) {
      throw new Error(`merge comparers must match join key count (${leftKeys.length}), got ${op.comparers.length}`);
    }
    keyIgnoreCase = op.comparers.map(isIgnoreCaseComparer);
  } else if (op.comparer) {
    const flag = isIgnoreCaseComparer(op.comparer);
    keyIgnoreCase = new Array(leftKeys.length).fill(flag);
  }

  const leftIndex = buildColumnIndex(left.columns);
  const rightIndex = buildColumnIndex(right.columns);
  const getLeftIndex = (name) => {
    const idx = leftIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${left.columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  };
  const getRightIndex = (name) => {
    const idx = rightIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${right.columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  };

  const leftKeyIdx = leftKeys.map(getLeftIndex);
  const rightKeyIdx = rightKeys.map(getRightIndex);

  /** @type {import("../../table.js").Column[]} */
  let outColumns = [];
  /** @type {string[] | null} */
  let nestedRightNames = null;
  /** @type {number[] | null} */
  let nestedRightIdx = null;
  /** @type {import("../../table.js").Column[] | null} */
  let nestedColumns = null;
  /** @type {Array<{ idx: number }>} */
  let rightColumnsToInclude = [];

  if (joinMode === "nested") {
    if (typeof op.newColumnName !== "string" || op.newColumnName.length === 0) {
      throw new Error("merge joinMode 'nested' requires newColumnName");
    }
    if (left.columns.some((c) => c.name === op.newColumnName)) {
      throw new Error(`Column '${op.newColumnName}' already exists`);
    }

    nestedRightNames = Array.isArray(op.rightColumns) ? op.rightColumns : right.columns.map((c) => c.name);
    nestedRightIdx = nestedRightNames.map(getRightIndex);
    nestedColumns = nestedRightNames.map((name, idx) => ({ name, type: right.columns[nestedRightIdx[idx]]?.type ?? "any" }));
    outColumns = [...left.columns, { name: op.newColumnName, type: "any" }];
  } else {
    const excludeRightKeys = new Set(rightKeys);
    rightColumnsToInclude = right.columns
      .map((_col, idx) => ({ idx }))
      .filter(({ idx }) => !excludeRightKeys.has(right.columns[idx].name));

    const rawOutNames = [...left.columns.map((c) => c.name), ...rightColumnsToInclude.map(({ idx }) => right.columns[idx].name)];
    const uniqueOutNames = makeUniqueColumnNames(rawOutNames);
    outColumns = [
      ...left.columns.map((col, idx) => ({ ...col, name: uniqueOutNames[idx] })),
      ...rightColumnsToInclude.map(({ idx }, outIdx) => ({ ...right.columns[idx], name: uniqueOutNames[left.columns.length + outIdx] })),
    ];
  }

  const rightKeyPrefix = makeSpillKeyPrefix("pq-join-right");

  /** @type {AsyncIterable<unknown[][]>} */
  const batches = (async function* () {
    /** @type {Map<string, unknown[][]> | null} */
    let inMemoryIndex = new Map();
    let rightRowCount = 0;
    let spilled = false;

    /**
     * @param {Map<string, unknown[][]>} map
     */
    const flushIndexToStore = async (map) => {
      for (const [key, rows] of map.entries()) {
        if (rows.length === 0) continue;
        const storeKey = `${rightKeyPrefix}:key:${key}`;
        for (let i = 0; i < rows.length; i += options.batchSize) {
          await options.store.putBatch(storeKey, rows.slice(i, i + options.batchSize));
        }
      }
    };

    try {
      // Build right-side index first.
      for await (const batch of right.batches) {
        throwIfAborted(options.signal);
        if (!Array.isArray(batch) || batch.length === 0) continue;

        /** @type {Map<string, unknown[][]>} */
        const grouped = new Map();

        for (const rowRaw of batch) {
          if (!Array.isArray(rowRaw)) continue;
          const row = rowRaw;
          const key = compositeKeyForRow(row, rightKeyIdx, keyIgnoreCase);
          rightRowCount += 1;

          if (inMemoryIndex) {
            const bucket = inMemoryIndex.get(key);
            if (bucket) bucket.push(row);
            else inMemoryIndex.set(key, [row]);

            if (rightRowCount >= options.maxInMemoryRows) {
              spilled = true;
              await flushIndexToStore(inMemoryIndex);
              inMemoryIndex = null;
            }
            continue;
          }

          const bucket = grouped.get(key);
          if (bucket) bucket.push(row);
          else grouped.set(key, [row]);
        }

        if (!inMemoryIndex) {
          for (const [key, rows] of grouped.entries()) {
            const storeKey = `${rightKeyPrefix}:key:${key}`;
            for (let i = 0; i < rows.length; i += options.batchSize) {
              await options.store.putBatch(storeKey, rows.slice(i, i + options.batchSize));
            }
          }
        }
      }

      if (spilled) {
        options.onProgress?.({
          type: "stream:spill",
          queryId: options.queryId ?? "<unknown>",
          operator: "merge",
          phase: "buildRight",
          rightRowCount,
        });
      }

      /** @type {unknown[][]} */
      let outBatch = [];

      /**
       * @param {unknown[]} row
       */
      const pushRow = (row) => {
        outBatch.push(row);
        if (outBatch.length >= options.batchSize) {
          const ready = outBatch;
          outBatch = [];
          return ready;
        }
        return null;
      };

      for await (const leftBatch of left.batches) {
        throwIfAborted(options.signal);
        if (!Array.isArray(leftBatch) || leftBatch.length === 0) continue;

        for (const leftRowRaw of leftBatch) {
          if (!Array.isArray(leftRowRaw)) continue;
          const leftRow = leftRowRaw;
          const normalizeCell = (value) => (value === undefined ? null : value);
          const leftValues = leftRow.map(normalizeCell);
          const key = compositeKeyForRow(leftRow, leftKeyIdx, keyIgnoreCase);

          if (joinMode === "nested") {
            /** @type {unknown[][]} */
            const matches = [];

            if (inMemoryIndex) {
              const rows = inMemoryIndex.get(key) ?? [];
              for (const rightRow of rows) matches.push(rightRow);
            } else {
              const storeKey = `${rightKeyPrefix}:key:${key}`;
              for await (const rightRow of options.store.iterateRows(storeKey)) {
                throwIfAborted(options.signal);
                matches.push(rightRow);
              }
            }

            if (matches.length === 0 && joinType === "inner") {
              continue;
            }

            const nestedRows =
              matches.length === 0
                ? []
                : matches.map((row) => nestedRightIdx.map((idx) => normalizeCell(row[idx])));
            const nestedTable = new DataTable(nestedColumns, nestedRows);
            {
              const ready = pushRow([...leftValues, nestedTable]);
              if (ready) yield ready;
            }
            continue;
          }

          // Flat join (Table.Join) semantics.
          let matched = false;

          if (inMemoryIndex) {
            const rows = inMemoryIndex.get(key) ?? [];
            for (const rightRow of rows) {
              matched = true;
              const out = [...leftValues, ...rightColumnsToInclude.map(({ idx }) => normalizeCell(rightRow[idx]))];
              const ready = pushRow(out);
              if (ready) yield ready;
            }
          } else {
            const storeKey = `${rightKeyPrefix}:key:${key}`;
            for await (const rightRow of options.store.iterateRows(storeKey)) {
              throwIfAborted(options.signal);
              matched = true;
              const out = [...leftValues, ...rightColumnsToInclude.map(({ idx }) => normalizeCell(rightRow[idx]))];
              const ready = pushRow(out);
              if (ready) yield ready;
            }
          }

          if (!matched && joinType === "left") {
            const out = [...leftValues, ...rightColumnsToInclude.map(() => null)];
            const ready = pushRow(out);
            if (ready) yield ready;
          }
        }
      }

      if (outBatch.length > 0) yield outBatch;
    } finally {
      try {
        await options.store.clearPrefix(rightKeyPrefix);
      } catch {
        // ignore
      }
      options.onProgress?.({
        type: "stream:operator",
        queryId: options.queryId ?? "<unknown>",
        operator: "merge",
        spilled,
      });
    }
  })();

  return { columns: outColumns, batches };
}
