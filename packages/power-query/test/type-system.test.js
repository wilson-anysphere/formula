import assert from "node:assert/strict";
import test from "node:test";

import { deserializeTable, serializeTable } from "../src/cache/serialize.js";
import { DataTable, inferColumnType } from "../src/table.js";
import { MS_PER_DAY, PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../src/values.js";

test("inferColumnType detects core Power Query scalar types", () => {
  assert.equal(inferColumnType([null, undefined]), "any");
  assert.equal(inferColumnType([new Date("2024-01-01T00:00:00.000Z")]), "date");
  assert.equal(inferColumnType([new Date("2024-01-01T12:34:56.000Z")]), "datetime");
  assert.equal(inferColumnType([new PqDateTimeZone(new Date("2024-01-01T00:00:00.000Z"), 0)]), "datetimezone");
  assert.equal(inferColumnType([new PqTime(0)]), "time");
  assert.equal(inferColumnType([new PqDuration(MS_PER_DAY)]), "duration");
  assert.equal(inferColumnType([new PqDecimal("1.25")]), "decimal");
  assert.equal(inferColumnType([new Uint8Array([1, 2, 3])]), "binary");
});

test("serializeTable/deserializeTable roundtrip expanded scalar types", () => {
  const dt = new Date("2024-01-02T03:04:05.678Z");
  const dtz = PqDateTimeZone.from("2024-01-02T03:04:05.678+02:00");
  assert.ok(dtz, "datetimezone literal should parse");
  const time = new PqTime(13 * 3_600_000 + 37 * 60_000 + 42 * 1000 + 250);
  const duration = new PqDuration(MS_PER_DAY + 1_234);
  const decimal = new PqDecimal("123.456");
  const binary = new Uint8Array([1, 2, 3]);

  const table = new DataTable(
    [
      { name: "dt", type: "datetime" },
      { name: "dtz", type: "datetimezone" },
      { name: "time", type: "time" },
      { name: "duration", type: "duration" },
      { name: "decimal", type: "decimal" },
      { name: "binary", type: "binary" },
    ],
    [[dt, dtz, time, duration, decimal, binary]],
  );

  const encoded = JSON.parse(JSON.stringify(serializeTable(table)));
  const restored = deserializeTable(encoded);

  assert.deepEqual(restored.columns, table.columns);

  const restoredGrid = restored.toGrid();
  assert.ok(restoredGrid[1][0] instanceof Date);
  assert.equal(restoredGrid[1][0].toISOString(), dt.toISOString());
  assert.ok(restoredGrid[1][1] instanceof PqDateTimeZone);
  assert.equal(restoredGrid[1][1].toString(), dtz.toString());
  assert.ok(restoredGrid[1][2] instanceof PqTime);
  assert.equal(restoredGrid[1][2].toString(), time.toString());
  assert.ok(restoredGrid[1][3] instanceof PqDuration);
  assert.equal(restoredGrid[1][3].toString(), duration.toString());
  assert.ok(restoredGrid[1][4] instanceof PqDecimal);
  assert.equal(restoredGrid[1][4].toString(), decimal.toString());
  assert.ok(restoredGrid[1][5] instanceof Uint8Array);
  assert.deepEqual(restoredGrid[1][5], binary);
});

