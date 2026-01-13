import assert from "node:assert/strict";
import test from "node:test";
import {
  ActionableError,
  filenameFromUrl,
  validateLatestJson,
} from "./verify-desktop-release-assets.mjs";

function assetMap(names) {
  return new Map(names.map((name) => [name, { name }]));
}

test("filenameFromUrl extracts decoded filename and strips query", () => {
  assert.equal(
    filenameFromUrl("https://example.com/download/My%20File.exe?foo=1#bar"),
    "My File.exe",
  );
});

test("validateLatestJson passes for a minimal manifest (version normalization + required OS keys)", () => {
  const manifest = {
    version: "v0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula_0.1.0_x86_64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula_0.1.0_x86_64.AppImage",
    "Formula_0.1.0_x64.msi",
    "Formula_0.1.0.app.tar.gz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson finds a nested platforms map", () => {
  const manifest = {
    version: "0.1.0",
    data: {
      platforms: {
        "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
        "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "sig" },
        "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      },
    },
  };

  const assets = assetMap(["Formula.AppImage", "Formula.msi", "Formula.app.tar.gz"]);
  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson fails when a required OS is missing", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "sig" },
    },
  };

  const assets = assetMap(["Formula.AppImage", "Formula.msi"]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing an entry containing \"darwin\""),
  );
});

test("validateLatestJson fails when a platform URL references a missing asset", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap(["Formula.AppImage", "Formula.msi"]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("but that asset is not present"),
  );
});

test("validateLatestJson accepts missing inline signature when a sibling .sig asset exists", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "" },
      "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula.AppImage.sig",
    "Formula.msi",
    "Formula.msi.sig",
    "Formula.app.tar.gz",
    "Formula.app.tar.gz.sig",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson fails when both inline signature and sibling .sig are missing", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "" },
      "windows-x86_64": { url: "https://example.com/Formula.msi", signature: "sig" },
      "darwin-universal": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap(["Formula.AppImage", "Formula.msi", "Formula.app.tar.gz"]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing a non-empty \"signature\""),
  );
});

