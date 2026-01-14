import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { findPlatformsObject, isMacUpdaterArchiveAssetName } from "./verify-tauri-updater-assets.mjs";
import { stripComments } from "../apps/desktop/test/sourceTextUtils.js";

test("isMacUpdaterArchiveAssetName matches macOS updater tarballs but rejects Linux AppImage tarballs", () => {
  assert.equal(isMacUpdaterArchiveAssetName("Formula.app.tar.gz"), true);
  assert.equal(isMacUpdaterArchiveAssetName("Formula.tar.gz"), true);
  assert.equal(isMacUpdaterArchiveAssetName("Formula.tgz"), true);

  assert.equal(isMacUpdaterArchiveAssetName("Formula.AppImage.tar.gz"), false);
  assert.equal(isMacUpdaterArchiveAssetName("Formula.AppImage.tgz"), false);
});

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

test("verify-tauri-updater-assets supports overriding tauri.conf.json path via FORMULA_TAURI_CONF_PATH", () => {
  const source = stripComments(readFileSync(new URL("./verify-tauri-updater-assets.mjs", import.meta.url), "utf8"));
  assert.match(source, /FORMULA_TAURI_CONF_PATH/);
});
