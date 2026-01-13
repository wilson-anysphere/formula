import assert from "node:assert/strict";
import test from "node:test";

import { validatePlatformEntries } from "./validate-updater-manifest.mjs";

function baseline() {
  const platforms = {
    "darwin-universal": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula.app.tar.gz",
      signature: "sig",
    },
    "windows-x86_64": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x64.msi",
      signature: "sig",
    },
    "windows-aarch64": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_arm64.msi",
      signature: "sig",
    },
    "linux-x86_64": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula.AppImage",
      signature: "sig",
    },
  };

  const assetNames = new Set([
    "Formula.app.tar.gz",
    "Formula_0.1.0_x64.msi",
    "Formula_0.1.0_arm64.msi",
    "Formula.AppImage",
  ]);

  return { platforms, assetNames };
}

test("fails when Linux updater URL points at a non-updatable package (.deb)", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/formula_0.1.0_amd64.deb";
  assetNames.delete("Formula.AppImage");
  assetNames.add("formula_0.1.0_amd64.deb");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.equal(result.invalidTargets.length, 0);
  assert.equal(result.missingAssets.length, 0);
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when Windows updater URL points at a non-updatable extension", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula.zip";
  assetNames.delete("Formula_0.1.0_x64.msi");
  assetNames.add("Formula.zip");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when two targets share the same URL (collision)", () => {
  const { platforms, assetNames } = baseline();
  const url = "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x64.msi";
  platforms["windows-x86_64"].url = url;
  platforms["windows-aarch64"].url = url;
  assetNames.delete("Formula_0.1.0_arm64.msi");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Duplicate platform URLs in latest.json")),
    `Expected duplicate URL error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when two targets reference the same asset name via different URLs (querystring collision)", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-aarch64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x64.msi?token=abc";
  assetNames.delete("Formula_0.1.0_arm64.msi");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Duplicate platform assets in latest.json")),
    `Expected duplicate asset-name error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("passes with distinct URLs and correct per-platform updater artifact types", () => {
  const { platforms, assetNames } = baseline();

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
  assert.equal(result.validatedTargets.length, 4);
});
