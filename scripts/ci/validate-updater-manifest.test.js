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
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_amd64.AppImage",
      signature: "sig",
    },
    "linux-aarch64": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_arm64.AppImage",
      signature: "sig",
    },
  };

  const assetNames = new Set([
    "Formula.app.tar.gz",
    "Formula_0.1.0_x64.msi",
    "Formula_0.1.0_arm64.msi",
    "Formula_0.1.0_amd64.AppImage",
    "Formula_0.1.0_arm64.AppImage",
  ]);

  return { platforms, assetNames };
}

test("fails when Linux updater URL points at a non-updatable package (.deb)", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/formula_0.1.0_amd64.deb";
  assetNames.delete("Formula_0.1.0_amd64.AppImage");
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
  assert.equal(result.validatedTargets.length, 5);
});

test("accepts a macOS updater archive ending with .tar.gz (not .app.tar.gz)", () => {
  const { platforms, assetNames } = baseline();
  platforms["darwin-universal"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula_universal.tar.gz";
  assetNames.delete("Formula.app.tar.gz");
  assetNames.add("Formula_universal.tar.gz");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
});

test("accepts Windows updater installers ending with .exe (NSIS strategy)", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x64.exe";
  platforms["windows-aarch64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_arm64.exe";
  assetNames.delete("Formula_0.1.0_x64.msi");
  assetNames.delete("Formula_0.1.0_arm64.msi");
  assetNames.add("Formula_0.1.0_x64.exe");
  assetNames.add("Formula_0.1.0_arm64.exe");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
});

test("fails when Windows updater assets do not include an arch token in the filename", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0.msi";
  assetNames.delete("Formula_0.1.0_x64.msi");
  assetNames.add("Formula_0.1.0.msi");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Invalid Windows updater asset naming in latest.json.platforms")),
    `Expected Windows arch-token validation error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when Linux updater assets do not include an arch token in the filename", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula.AppImage";
  assetNames.delete("Formula_0.1.0_amd64.AppImage");
  assetNames.add("Formula.AppImage");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Invalid Linux updater asset naming in latest.json.platforms")),
    `Expected Linux arch-token validation error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when latest.json.platforms is missing a required platform key", () => {
  const { platforms, assetNames } = baseline();
  delete platforms["linux-x86_64"];
  assetNames.delete("Formula_0.1.0_amd64.AppImage");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Unexpected latest.json.platforms keys")),
    `Expected strict platforms key mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when latest.json.platforms contains an unexpected platform key", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-i686"] = {
    url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x86.msi",
    signature: "sig",
  };
  assetNames.add("Formula_0.1.0_x86.msi");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Unexpected latest.json.platforms keys")),
    `Expected strict platforms key mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("reports missing release assets referenced by platforms[*].url", () => {
  const { platforms, assetNames } = baseline();
  // Keep the file extension valid, but remove it from the release asset set.
  assetNames.delete("Formula_0.1.0_x64.msi");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.equal(result.missingAssets.length, 1);
  assert.equal(result.missingAssets[0].target, "windows-x86_64");
  assert.equal(result.missingAssets[0].assetName, "Formula_0.1.0_x64.msi");
});
