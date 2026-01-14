import assert from "node:assert/strict";
import crypto from "node:crypto";
import { readFileSync } from "node:fs";
import test from "node:test";
import {
  ActionableError,
  filenameFromUrl,
  isPrimaryBundleAssetName,
  isPrimaryBundleOrSig,
  validateReleaseExpectations,
  validateLatestJson,
  verifyUpdaterManifestSignature,
} from "./verify-desktop-release-assets.mjs";

import { stripComments } from "../apps/desktop/test/sourceTextUtils.js";

function assetMap(names) {
  return new Map(names.map((name) => [name, { name }]));
}

test("filenameFromUrl extracts decoded filename and strips query", () => {
  assert.equal(
    filenameFromUrl("https://example.com/download/My%20File.exe?foo=1#bar"),
    "My File.exe",
  );
});

test("verify-desktop-release-assets supports overriding tauri.conf.json path via FORMULA_TAURI_CONF_PATH", () => {
  const source = stripComments(readFileSync(new URL("./verify-desktop-release-assets.mjs", import.meta.url), "utf8"));
  assert.match(source, /FORMULA_TAURI_CONF_PATH/);
});

test("isPrimaryBundleAssetName matches expected suffixes", () => {
  assert.equal(isPrimaryBundleAssetName("Formula.dmg"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.app.tar.gz"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.tar.gz"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.tgz"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.msi"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.exe"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.AppImage"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.deb"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.rpm"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.zip"), true);
  assert.equal(isPrimaryBundleAssetName("Formula.pkg"), true);
  assert.equal(isPrimaryBundleAssetName("latest.json"), false);
});

test("isPrimaryBundleOrSig includes .sig when enabled", () => {
  assert.equal(isPrimaryBundleOrSig("Formula.msi", { includeSigs: false }), true);
  assert.equal(isPrimaryBundleOrSig("Formula.msi.sig", { includeSigs: false }), false);
  assert.equal(isPrimaryBundleOrSig("Formula.msi.sig", { includeSigs: true }), true);
  assert.equal(isPrimaryBundleOrSig("latest.json.sig", { includeSigs: true }), false);
});

test("validateLatestJson passes for a minimal manifest (version normalization + required OS keys)", () => {
  const manifest = {
    version: "v0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula_0.1.0_x86_64.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
      // macOS universal builds are published as a single updater archive but are referenced by
      // both macOS arch updater keys.
      "darwin-x86_64": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula_0.1.0_x86_64.AppImage",
    "Formula_0.1.0_arm64.AppImage",
    "Formula_0.1.0_x64.msi",
    "Formula_0.1.0_arm64.msi",
    "Formula_0.1.0.app.tar.gz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson accepts macOS updater tarballs (.app.tar.gz preferred; allow .tar.gz/.tgz)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.tgz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.tar.gz",
    "Formula.tgz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson rejects Linux AppImage tarballs for macOS updater keys", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.AppImage.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.AppImage.tar.gz",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("darwin-x86_64") &&
      /appimage tarballs?/i.test(err.message),
  );
});

test("validateLatestJson allows installer-specific platform keys to reference installer artifacts", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      // Some Tauri versions may emit additional `{os}-{arch}-{bundle}` keys. These should still
      // reference existing assets, but are not the runtime updater targets and may legitimately point
      // at installer bundles like `.deb`/`.rpm`/`.dmg`.
      "linux-x86_64-deb": { url: "https://example.com/Formula.deb", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
    "Formula.deb",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson finds a nested platforms map", () => {
  const manifest = {
    version: "0.1.0",
    data: {
      platforms: {
        "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
        "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
        "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
        "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
        "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
        "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);
  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson fails when a required OS is missing", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      // Include one macOS entry, but intentionally omit the other to exercise required-key validation.
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing required key \"darwin-x86_64\""),
  );
});

test("validateLatestJson fails when a platform URL references a missing asset", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
  ]);
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
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula.AppImage.sig",
    "Formula_arm64.AppImage",
    "Formula_arm64.AppImage.sig",
    "Formula_x64.msi",
    "Formula_x64.msi.sig",
    "Formula_arm64.msi",
    "Formula_arm64.msi.sig",
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
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);
  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) => err instanceof ActionableError && err.message.includes("missing a non-empty \"signature\""),
  );
});

test("validateLatestJson rejects macOS .dmg updater URLs (even if asset exists)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.dmg", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.dmg",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("darwin-x86_64") &&
      err.message.includes("Formula.dmg") &&
      err.message.includes("Expected file extensions"),
  );
});

test("validateLatestJson accepts a macOS updater archive ending with .tgz", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      // macOS universal builds may be published as .tgz; the updater keys still reference that tarball.
      "darwin-x86_64": { url: "https://example.com/Formula_universal.tgz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula_universal.tgz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula_universal.tgz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("validateLatestJson rejects macOS .pkg updater URLs (even if asset exists)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.pkg", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.pkg",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("darwin-x86_64") &&
      err.message.includes("Formula.pkg") &&
      err.message.includes("Expected file extensions"),
  );
});

test("validateLatestJson rejects macOS updater keys pointing at Linux AppImage tarballs (.AppImage.tgz)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      // This would otherwise match the macOS \"*.tar.gz\" rule, but should be rejected.
      "darwin-x86_64": { url: "https://example.com/Formula.AppImage.tgz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.AppImage.tgz",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("darwin-x86_64") &&
      err.message.includes("Formula.AppImage.tgz") &&
      /appimage tarballs?/i.test(err.message),
  );
});

test("validateLatestJson rejects Linux .deb/.rpm updater URLs (even if asset exists)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.deb", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.deb",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("linux-x86_64") &&
      err.message.includes("Formula.deb"),
  );
});

test("validateLatestJson rejects Linux .rpm updater URLs (even if asset exists)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.rpm", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.rpm",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("linux-x86_64") &&
      err.message.includes("Formula.rpm"),
  );
});

test("validateLatestJson rejects Windows .zip updater URLs (even if asset exists)", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi.zip", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi.zip",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("windows-x86_64") &&
      err.message.includes(".zip archive"),
  );
});

test("validateLatestJson rejects raw Windows .exe updater URLs by default", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.exe", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.exe",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);

  assert.throws(
    () => validateLatestJson(manifest, "0.1.0", assets),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("windows-x86_64") &&
      err.message.includes("--allow-windows-exe"),
  );
});

test("validateLatestJson allows raw Windows .exe updater URLs when explicitly enabled", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.exe", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.exe",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets, { allowWindowsExe: true }));
});

test("validateLatestJson allows installer-specific platform keys to reference installers (.exe) without affecting required updater keys", () => {
  const manifest = {
    version: "0.1.0",
    platforms: {
      "linux-x86_64": { url: "https://example.com/Formula.AppImage", signature: "sig" },
      "linux-aarch64": { url: "https://example.com/Formula_arm64.AppImage", signature: "sig" },
      "windows-x86_64": { url: "https://example.com/Formula_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_arm64.msi", signature: "sig" },
      "darwin-x86_64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula.app.tar.gz", signature: "sig" },
      // Additional (installer-specific) key: allowed to reference an installer like NSIS `.exe`.
      "windows-x86_64-nsis": { url: "https://example.com/Formula_x64.exe", signature: "sig" },
    },
  };

  const assets = assetMap([
    "Formula.AppImage",
    "Formula_arm64.AppImage",
    "Formula_x64.msi",
    "Formula_arm64.msi",
    "Formula.app.tar.gz",
    "Formula_x64.exe",
  ]);

  assert.doesNotThrow(() => validateLatestJson(manifest, "0.1.0", assets));
});

test("verifyUpdaterManifestSignature verifies latest.json.sig against latest.json with minisign pubkey", () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }) + "\n", "utf8");
  const signature = crypto.sign(null, latestJsonBytes, privateKey);
  const latestSigText = `${signature.toString("base64")}\n`;

  // Build a Tauri/minisign pubkey string: base64(minisign text block).
  const spki = /** @type {Buffer} */ (publicKey.export({ format: "der", type: "spki" }));
  const rawPubkey = spki.subarray(spki.length - 32);
  const keyIdLe = Buffer.from("0102030405060708", "hex");
  const keyIdHex = Buffer.from(keyIdLe).reverse().toString("hex").toUpperCase();
  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), keyIdLe, rawPubkey]);
  const pubkeyText = `untrusted comment: minisign public key: ${keyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  assert.doesNotThrow(() => verifyUpdaterManifestSignature(latestJsonBytes, latestSigText, pubkeyBase64));
});

test("verifyUpdaterManifestSignature fails on key id mismatch (minisign payload)", () => {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const latestJsonBytes = Buffer.from(JSON.stringify({ version: "0.1.0", platforms: {} }) + "\n", "utf8");
  const signature = crypto.sign(null, latestJsonBytes, privateKey);

  const spki = /** @type {Buffer} */ (publicKey.export({ format: "der", type: "spki" }));
  const rawPubkey = spki.subarray(spki.length - 32);

  const pubKeyIdLe = Buffer.from("0102030405060708", "hex");
  const pubKeyIdHex = Buffer.from(pubKeyIdLe).reverse().toString("hex").toUpperCase();
  const pubPayload = Buffer.concat([Buffer.from([0x45, 0x64]), pubKeyIdLe, rawPubkey]);
  const pubkeyText = `untrusted comment: minisign public key: ${pubKeyIdHex}\n${pubPayload.toString("base64")}\n`;
  const pubkeyBase64 = Buffer.from(pubkeyText, "utf8").toString("base64");

  const sigKeyIdLe = Buffer.from("1111111111111111", "hex");
  const minisignSigPayload = Buffer.concat([Buffer.from([0x45, 0x64]), sigKeyIdLe, Buffer.from(signature)]);
  const latestSigText = `${minisignSigPayload.toString("base64")}\n`;

  assert.throws(
    () => verifyUpdaterManifestSignature(latestJsonBytes, latestSigText, pubkeyBase64),
    (err) => err instanceof ActionableError && /key id mismatch/i.test(err.message),
  );
});

test("validateReleaseExpectations passes for complete multi-arch windows assets + updater keys", () => {
  const expectedTargets = [
    {
      id: "windows-x64",
      os: "windows",
      arch: "x64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-x86_64", "windows-x64"],
    },
    {
      id: "windows-arm64",
      os: "windows",
      arch: "arm64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-aarch64", "windows-arm64"],
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
    },
  };

  const assetNames = ["Formula_0.1.0_x64.msi", "Formula_0.1.0_arm64.msi"];

  assert.doesNotThrow(() =>
    validateReleaseExpectations({
      manifest,
      expectedVersion: "0.1.0",
      assetNames,
      expectedTargets,
    }),
  );
});

test("validateReleaseExpectations fails when an expected installer is missing", () => {
  const expectedTargets = [
    {
      id: "windows-x64",
      os: "windows",
      arch: "x64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-x86_64"],
    },
    {
      id: "windows-arm64",
      os: "windows",
      arch: "arm64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-aarch64"],
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
    },
  };

  // Missing the arm64 installer.
  const assetNames = ["Formula_0.1.0_x64.msi"];

  assert.throws(
    () =>
      validateReleaseExpectations({
        manifest,
        expectedVersion: "0.1.0",
        assetNames,
        expectedTargets,
      }),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("expectation errors") &&
      err.message.includes("Missing installer asset"),
  );
});

test("validateReleaseExpectations fails when latest.json is missing an expected updater platform key", () => {
  const expectedTargets = [
    {
      id: "windows-x64",
      os: "windows",
      arch: "x64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-x86_64"],
    },
    {
      id: "windows-arm64",
      os: "windows",
      arch: "arm64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-aarch64"],
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      // Missing windows-aarch64
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
    },
  };

  const assetNames = ["Formula_0.1.0_x64.msi", "Formula_0.1.0_arm64.msi"];

  assert.throws(
    () =>
      validateReleaseExpectations({
        manifest,
        expectedVersion: "0.1.0",
        assetNames,
        expectedTargets,
      }),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("Missing updater platform entry") &&
      err.message.includes("windows-arm64"),
  );
});

test("validateReleaseExpectations fails on ambiguous multi-arch assets that omit arch tokens", () => {
  const expectedTargets = [
    {
      id: "windows-x64",
      os: "windows",
      arch: "x64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-x86_64"],
    },
    {
      id: "windows-arm64",
      os: "windows",
      arch: "arm64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-aarch64"],
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
    },
  };

  // Ambiguous installer: would collide across architectures if uploaded twice.
  const assetNames = ["Formula_0.1.0_x64.msi", "Formula_0.1.0_arm64.msi", "Formula_0.1.0.msi"];

  assert.throws(
    () =>
      validateReleaseExpectations({
        manifest,
        expectedVersion: "0.1.0",
        assetNames,
        expectedTargets,
      }),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("Ambiguous artifacts") &&
      err.message.includes("windows"),
  );
});

test("validateReleaseExpectations fails when an asset matches multiple arch tokens (e.g. x64+arm64)", () => {
  const expectedTargets = [
    {
      id: "windows-x64",
      os: "windows",
      arch: "x64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-x86_64"],
    },
    {
      id: "windows-arm64",
      os: "windows",
      arch: "arm64",
      installerExts: [".msi", ".exe"],
      updaterPlatformKeys: ["windows-aarch64"],
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      "windows-x86_64": { url: "https://example.com/Formula_0.1.0_x64.msi", signature: "sig" },
      "windows-aarch64": { url: "https://example.com/Formula_0.1.0_arm64.msi", signature: "sig" },
    },
  };

  // Single installer name includes both tokens; this should not be accepted as satisfying both
  // arch expectations.
  const assetNames = ["Formula_0.1.0_x64_arm64.msi"];

  assert.throws(
    () =>
      validateReleaseExpectations({
        manifest,
        expectedVersion: "0.1.0",
        assetNames,
        expectedTargets,
      }),
    (err) =>
      err instanceof ActionableError &&
      err.message.includes("multiple architecture tokens") &&
      err.message.includes("windows"),
  );
});

test("validateReleaseExpectations allows missing arch token for macos-universal installers", () => {
  const expectedTargets = [
    {
      id: "macos-universal",
      os: "macos",
      arch: "universal",
      installerExts: [".dmg", ".pkg"],
      updaterPlatformKeys: ["darwin-x86_64", "darwin-aarch64"],
      allowMissingArchInInstallerName: true,
    },
  ];

  const manifest = {
    version: "0.1.0",
    platforms: {
      "darwin-x86_64": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
      "darwin-aarch64": { url: "https://example.com/Formula_0.1.0.app.tar.gz", signature: "sig" },
    },
  };

  const assetNames = ["Formula_0.1.0.dmg", "Formula_0.1.0.app.tar.gz"];

  assert.doesNotThrow(() =>
    validateReleaseExpectations({
      manifest,
      expectedVersion: "0.1.0",
      assetNames,
      expectedTargets,
    }),
  );
});
