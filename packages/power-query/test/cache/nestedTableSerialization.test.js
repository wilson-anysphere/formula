import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { QueryEngine } from "../../src/engine.js";
import { compileMToQuery } from "../../src/m/compiler.js";
import { DataTable } from "../../src/table.js";

test("QueryEngine: cached Table.NestedJoin results preserve nested table cells", async () => {
  const cache = new CacheManager({ store: new MemoryCacheStore() });
  const engine = new QueryEngine({ cache });

  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id"},
    {1},
    {2}
  })
in
  Source
`,
    { id: "q_left", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Value"},
    {2, "B"}
  })
in
  Source
`,
    { id: "q_right", name: "Right" },
  );

  const nestedJoin = compileMToQuery(
    `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.NestedJoin(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter)
in
  #"Merged Queries"
`,
    { id: "q_nested", name: "Nested join" },
  );

  const context = { queries: { q_left: left, q_right: right } };

  const first = await engine.executeQueryWithMeta(nestedJoin, context, {});
  assert.equal(first.meta.cache?.hit, false);

  const matchesIdx = first.table.getColumnIndex("Matches");
  const firstRow0 = first.table.getCell(0, matchesIdx);
  assert.ok(firstRow0 instanceof DataTable);
  assert.deepEqual(firstRow0.toGrid(), [["Id", "Value"]]);

  const firstRow1 = first.table.getCell(1, matchesIdx);
  assert.ok(firstRow1 instanceof DataTable);
  assert.deepEqual(firstRow1.toGrid(), [
    ["Id", "Value"],
    [2, "B"],
  ]);

  const second = await engine.executeQueryWithMeta(nestedJoin, context, {});
  assert.equal(second.meta.cache?.hit, true);

  const matchesIdx2 = second.table.getColumnIndex("Matches");
  const secondRow0 = second.table.getCell(0, matchesIdx2);
  assert.ok(secondRow0 instanceof DataTable, "cached nested table cell should deserialize back into a DataTable");
  assert.deepEqual(secondRow0.toGrid(), [["Id", "Value"]]);

  const secondRow1 = second.table.getCell(1, matchesIdx2);
  assert.ok(secondRow1 instanceof DataTable);
  assert.deepEqual(secondRow1.toGrid(), [
    ["Id", "Value"],
    [2, "B"],
  ]);
});

