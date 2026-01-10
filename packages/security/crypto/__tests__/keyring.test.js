import test from "node:test";
import assert from "node:assert/strict";

import { KeyRing } from "../keyring.js";

test("KeyRing encrypt/decrypt across rotations", () => {
  const ring = KeyRing.create();
  const aadContext = { docId: "doc-1" };

  const first = ring.encrypt(Buffer.from("v1", "utf8"), { aadContext });
  ring.rotate();
  const second = ring.encrypt(Buffer.from("v2", "utf8"), { aadContext });

  assert.equal(ring.decrypt(first, { aadContext }).toString("utf8"), "v1");
  assert.equal(ring.decrypt(second, { aadContext }).toString("utf8"), "v2");
});

test("KeyRing JSON round-trip preserves old keys", () => {
  const ring = KeyRing.create();
  const first = ring.encrypt(Buffer.from("hello", "utf8"));
  ring.rotate();

  const restored = KeyRing.fromJSON(ring.toJSON());
  assert.equal(restored.decrypt(first).toString("utf8"), "hello");
});

