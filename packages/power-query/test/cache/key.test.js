import assert from "node:assert/strict";
import test from "node:test";

import { stableStringify } from "../../src/cache/key.js";
import { PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../../src/values.js";

test("stableStringify tags Power Query wrapper values to avoid cache key collisions", () => {
  const time = new PqTime(0);
  const duration = new PqDuration(0);
  assert.notEqual(stableStringify(time), stableStringify(duration));

  assert.deepEqual(JSON.parse(stableStringify(time)), { $type: "time", value: "00:00:00" });
  assert.deepEqual(JSON.parse(stableStringify(duration)), { $type: "duration", value: "PT0S" });

  const dtz = new PqDateTimeZone(new Date("2024-01-01T00:00:00.000Z"), 0);
  assert.deepEqual(JSON.parse(stableStringify(dtz)), { $type: "datetimezone", value: "2024-01-01T00:00:00.000Z" });

  const decimal = new PqDecimal("123.450");
  assert.deepEqual(JSON.parse(stableStringify(decimal)), { $type: "decimal", value: "123.450" });
});

test("stableStringify encodes Uint8Array values as base64", () => {
  const bytes = new Uint8Array([1, 2, 3]);
  assert.deepEqual(JSON.parse(stableStringify(bytes)), { $type: "binary", value: "AQID" });
});

