import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { validatePlatformEntries } from "./validate-updater-manifest.mjs";
import { stripComments } from "../../apps/desktop/test/sourceTextUtils.js";

function baseline() {
  const platforms = {
    "darwin-x86_64": {
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula.app.tar.gz",
      signature: "sig",
    },
    "darwin-aarch64": {
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
      url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x86_64.AppImage",
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
    "Formula_0.1.0_x86_64.AppImage",
    "Formula_0.1.0_arm64.AppImage",
  ]);

  return { platforms, assetNames };
}

test("fails when Linux updater URL points at a non-updatable package (.deb)", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/formula_0.1.0_amd64.deb";
  assetNames.delete("Formula_0.1.0_x86_64.AppImage");
  assetNames.add("formula_0.1.0_amd64.deb");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.equal(result.invalidTargets.length, 0);
  assert.equal(result.missingAssets.length, 0);
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when Linux updater URL points at a non-updatable package (.rpm)", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/formula_0.1.0_amd64.rpm";
  assetNames.delete("Formula_0.1.0_x86_64.AppImage");
  assetNames.add("formula_0.1.0_amd64.rpm");

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
    result.errors.some((e) => e.includes("Duplicate updater URLs across required targets")),
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
    result.errors.some((e) => e.includes("Duplicate updater assets across required targets")),
    `Expected duplicate asset-name error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("passes with distinct URLs and correct per-platform updater artifact types", () => {
  const { platforms, assetNames } = baseline();

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
  assert.equal(result.validatedTargets.length, 6);
});

test("validate-updater-manifest supports overriding tauri.conf.json path via FORMULA_TAURI_CONF_PATH", () => {
  const source = stripComments(readFileSync(new URL("./validate-updater-manifest.mjs", import.meta.url), "utf8"));
  assert.match(source, /FORMULA_TAURI_CONF_PATH/);
});

test("passes when installer-specific platform keys are present (e.g. linux-x86_64-deb)", () => {
  const { platforms, assetNames } = baseline();
  platforms["linux-x86_64-deb"] = {
    url: "https://github.com/example/repo/releases/download/v0.1.0/formula_0.1.0_amd64.deb",
    signature: "sig",
  };
  assetNames.add("formula_0.1.0_amd64.deb");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
});

test("fails when a macOS updater entry points at a non-updater artifact (.dmg)", () => {
  const { platforms, assetNames } = baseline();
  platforms["darwin-x86_64"].url =
    "https://github.com/example/repo/releases/download/v0.1.0/Formula.dmg";
  assetNames.delete("Formula.app.tar.gz");
  assetNames.add("Formula.dmg");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected macOS asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("accepts macOS updater tarballs that are not .app.tar.gz (allow .tar.gz/.tgz)", () => {
  const { platforms, assetNames } = baseline();
  const url = "https://github.com/example/repo/releases/download/v0.1.0/Formula_universal.tar.gz";
  platforms["darwin-x86_64"].url = url;
  platforms["darwin-aarch64"].url = url;
  assetNames.delete("Formula.app.tar.gz");
  assetNames.add("Formula_universal.tar.gz");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
});

test("accepts macOS updater tarballs ending with .tgz", () => {
  const { platforms, assetNames } = baseline();
  const url = "https://github.com/example/repo/releases/download/v0.1.0/Formula_universal.tgz";
  platforms["darwin-x86_64"].url = url;
  platforms["darwin-aarch64"].url = url;
  assetNames.delete("Formula.app.tar.gz");
  assetNames.add("Formula_universal.tgz");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
});

test("fails when a macOS updater entry points at a Linux AppImage tarball", () => {
  const { platforms, assetNames } = baseline();
  const url = "https://github.com/example/repo/releases/download/v0.1.0/Formula.AppImage.tar.gz";
  platforms["darwin-x86_64"].url = url;
  platforms["darwin-aarch64"].url = url;
  assetNames.delete("Formula.app.tar.gz");
  assetNames.add("Formula.AppImage.tar.gz");

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected macOS asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
});

test("fails when Windows updater installer is .exe instead of the expected .msi", () => {
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
  assert.ok(
    result.errors.some((e) => e.includes("Updater asset type mismatch in latest.json.platforms")),
    `Expected Windows asset type mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
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
  assetNames.delete("Formula_0.1.0_x86_64.AppImage");
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
  assetNames.delete("Formula_0.1.0_x86_64.AppImage");

  const result = validatePlatformEntries({ platforms, assetNames });
  const message = result.errors.join("\n");
  assert.ok(
    result.errors.some((e) => e.includes("Missing required latest.json.platforms keys")),
    `Expected strict platforms key mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
  assert.match(message, /Missing \(1\):/);
  assert.match(message, /linux-x86_64/);
});

test("fails when latest.json.platforms is missing linux-aarch64", () => {
  const { platforms, assetNames } = baseline();
  delete platforms["linux-aarch64"];
  assetNames.delete("Formula_0.1.0_arm64.AppImage");

  const result = validatePlatformEntries({ platforms, assetNames });
  const message = result.errors.join("\n");
  assert.ok(
    result.errors.some((e) => e.includes("Missing required latest.json.platforms keys")),
    `Expected strict platforms key mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
  assert.match(message, /Missing \(1\):/);
  assert.match(message, /linux-aarch64/);
});

test("fails when latest.json.platforms is missing darwin-aarch64 (macOS universal per-arch key)", () => {
  const { platforms, assetNames } = baseline();
  delete platforms["darwin-aarch64"];

  const result = validatePlatformEntries({ platforms, assetNames });
  const message = result.errors.join("\n");
  assert.ok(
    result.errors.some((e) => e.includes("Missing required latest.json.platforms keys")),
    `Expected strict platforms key mismatch error, got:\n${result.errors.join("\n\n")}`,
  );
  assert.match(message, /darwin-aarch64/);
});

test("accepts additional installer-specific platform keys (e.g. windows-x86_64-msi)", () => {
  const { platforms, assetNames } = baseline();
  platforms["windows-x86_64-msi"] = {
    url: "https://github.com/example/repo/releases/download/v0.1.0/Formula_0.1.0_x64.msi",
    signature: "sig",
  };

  const result = validatePlatformEntries({ platforms, assetNames });
  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.invalidTargets, []);
  assert.deepEqual(result.missingAssets, []);
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
