import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";

function collectBatches(batches) {
  const grid = [];
  for (const batch of batches) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }
  return grid;
}

const leftCsv = ["Id,Val", "1,a", "2,b", "2,c", "3,d"].join("\n") + "\n";
const rightCsv = ["Id,Other", "2,10", "2,20", "4,30"].join("\n") + "\n";

/**
 * @param {string} path
 */
function csvForPath(path) {
  if (path.includes("left")) return leftCsv;
  if (path.includes("right")) return rightCsv;
  throw new Error(`unexpected path ${path}`);
}

test("streaming v2 merge (inner/left) matches materialized execution", async () => {
  const engineStreaming = new QueryEngine({
    fileAdapter: {
      readText: async () => {
        throw new Error("readText should not be called in streaming mode");
      },
      readTextStream: async function* (path) {
        yield csvForPath(path);
      },
    },
  });

  const engineMaterialized = new QueryEngine({
    fileAdapter: {
      readText: async (path) => csvForPath(path),
    },
  });

  const leftQuery = {
    id: "q_left",
    name: "Left",
    source: { type: "csv", path: "/tmp/left.csv", options: { hasHeaders: true } },
    steps: [],
  };

  const rightQuery = {
    id: "q_right",
    name: "Right",
    source: { type: "csv", path: "/tmp/right.csv", options: { hasHeaders: true } },
    steps: [],
  };

  const joinLeft = {
    id: "q_join_left",
    name: "Join left",
    source: { type: "query", queryId: "q_left" },
    steps: [
      {
        id: "s_join",
        name: "Join",
        operation: { type: "merge", rightQuery: "q_right", leftKeys: ["Id"], rightKeys: ["Id"], joinType: "left" },
      },
    ],
  };

  const joinInnerNestedExpand = {
    id: "q_join_nested_expand",
    name: "Join nested + expand",
    source: { type: "query", queryId: "q_left" },
    steps: [
      {
        id: "s_join",
        name: "Join",
        operation: {
          type: "merge",
          rightQuery: "q_right",
          leftKeys: ["Id"],
          rightKeys: ["Id"],
          joinType: "left",
          joinMode: "nested",
          newColumnName: "Matches",
          rightColumns: ["Other"],
        },
      },
      {
        id: "s_expand",
        name: "Expand",
        operation: { type: "expandTableColumn", column: "Matches", columns: ["Other"], newColumnNames: ["Other"] },
      },
    ],
  };

  for (const query of [joinLeft, joinInnerNestedExpand]) {
    const batches = [];
    await engineStreaming.executeQueryStreaming(query, { queries: { q_left: leftQuery, q_right: rightQuery } }, {
      batchSize: 10,
      materialize: false,
      onBatch: (batch) => batches.push(batch),
      streaming: { spill: { kind: "memory" }, maxInMemoryRows: 2 },
    });

    const streamed = collectBatches(batches);
    const expected = (await engineMaterialized.executeQuery(query, { queries: { q_left: leftQuery, q_right: rightQuery } }, {})).toGrid();
    assert.deepEqual(streamed, expected);
  }
});

