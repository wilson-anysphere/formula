import assert from "node:assert/strict";
import test from "node:test";

import { KeyRing } from "../../../packages/security/crypto/keyring.js";

import {
  generateKeyRingJson,
  rotateKeyRingJson,
  validateKeyRingJson,
} from "../src/keyring-cli.js";

test("keyring generate produces JSON parseable by KeyRing.fromJSON()", () => {
  const json = generateKeyRingJson();
  const ring = KeyRing.fromJSON(JSON.parse(JSON.stringify(json)));
  assert.equal(ring.currentVersion, 1);
  assert.deepEqual([...ring.keysByVersion.keys()], [1]);
});

test("keyring rotate increments currentVersion and preserves previous keys", () => {
  const original = generateKeyRingJson();
  const rotated = rotateKeyRingJson(original);

  const ring = KeyRing.fromJSON(rotated);
  assert.equal(ring.currentVersion, 2);

  const versions = [...ring.keysByVersion.keys()].sort((a, b) => a - b);
  assert.deepEqual(versions, [1, 2]);

  assert.equal(rotated.keys["1"], original.keys["1"]);
  assert.ok(typeof rotated.keys["2"] === "string" && rotated.keys["2"].length > 0);
});

test("keyring validate reports current and available versions", () => {
  const json = generateKeyRingJson();
  const summary = validateKeyRingJson(json);
  assert.deepEqual(summary, { currentVersion: 1, availableVersions: [1] });
});
