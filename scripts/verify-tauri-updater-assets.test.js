import assert from "node:assert/strict";
import test from "node:test";

import { findPlatformsObject } from "./verify-tauri-updater-assets.mjs";

test("findPlatformsObject finds top-level platforms", () => {
  const result = findPlatformsObject({
    version: "0.0.0",
    platforms: { "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" } },
  });

  assert.ok(result);
  assert.deepEqual(result.path, ["platforms"]);
  assert.ok("darwin-x86_64" in result.platforms);
});

test("findPlatformsObject finds nested platforms", () => {
  const result = findPlatformsObject({
    meta: { something: true },
    data: {
      platforms: {
        "windows-x86_64": { url: "https://example.com/app.msi", signature: "sig" },
      },
    },
  });

  assert.ok(result);
  assert.deepEqual(result.path, ["data", "platforms"]);
  assert.ok("windows-x86_64" in result.platforms);
});

test("findPlatformsObject returns null when missing", () => {
  assert.equal(findPlatformsObject({ version: "0.0.0" }), null);
  assert.equal(findPlatformsObject(null), null);
});
