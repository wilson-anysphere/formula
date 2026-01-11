import assert from "node:assert/strict";
import { readFileSync, readdirSync } from "node:fs";
import test from "node:test";

import { compileMToQuery } from "../../src/m/compiler.js";

const goldenDir = new URL("./golden/", import.meta.url);

/**
 * Convert Dates and other non-JSON values into a snapshot-friendly shape.
 * @param {unknown} value
 * @returns {any}
 */
function toSnapshotJson(value) {
  return JSON.parse(JSON.stringify(value));
}

for (const file of readdirSync(goldenDir).filter((f) => f.endsWith(".m")).sort()) {
  test(`m_language golden: ${file}`, () => {
    const mPath = new URL(file, goldenDir);
    const jsonPath = new URL(file.replace(/\.m$/, ".json"), goldenDir);
    const source = readFileSync(mPath, "utf8");
    const expected = JSON.parse(readFileSync(jsonPath, "utf8"));
    const actual = toSnapshotJson(compileMToQuery(source));
    assert.deepEqual(actual, expected);
  });
}
