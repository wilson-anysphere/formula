import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readSnapshot(snapshotName) {
  const filePath = path.join(__dirname, "__snapshots__", snapshotName);
  return fs.readFileSync(filePath, "utf8");
}

export function expectSnapshot(snapshotName, received) {
  const expected = readSnapshot(snapshotName);
  assert.equal(received, expected);
}
