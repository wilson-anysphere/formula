import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "generate-release-checksums.sh");

function touch(filePath) {
  mkdirSync(path.dirname(filePath), { recursive: true });
  writeFileSync(filePath, "");
}

/**
 * @param {string} assetsDir
 * @param {string} outputPath
 */
function runGenerateChecksums(assetsDir, outputPath) {
  return spawnSync("bash", [scriptPath, assetsDir, outputPath], {
    encoding: "utf8",
  });
}

/**
 * Minimal set of files the script expects, plus a macOS updater tarball.
 * @param {string} dir
 * @param {{ macUpdaterTarball: string }} opts
 */
function writeFixture(dir, { macUpdaterTarball }) {
  touch(path.join(dir, "Formula.dmg"));
  touch(path.join(dir, macUpdaterTarball));

  touch(path.join(dir, "Formula_x64.msi"));
  touch(path.join(dir, "Formula_x64.exe"));

  touch(path.join(dir, "Formula_x86_64.AppImage"));
  touch(path.join(dir, "Formula_amd64.deb"));
  touch(path.join(dir, "Formula_amd64.rpm"));

  touch(path.join(dir, "latest.json"));
  touch(path.join(dir, "latest.json.sig"));
}

test("generate-release-checksums accepts a standard macOS updater archive (*.app.tar.gz)", { skip: process.platform === "win32" }, () => {
  const tmp = mkdtempSync(path.join(os.tmpdir(), "formula-checksums-"));
  try {
    writeFixture(tmp, { macUpdaterTarball: "Formula.app.tar.gz" });
    const outPath = path.join(tmp, "SHA256SUMS.txt");
    const proc = runGenerateChecksums(tmp, outPath);
    assert.equal(proc.status, 0, proc.stderr);
  } finally {
    rmSync(tmp, { recursive: true, force: true });
  }
});

test("generate-release-checksums rejects Linux AppImage tarballs as macOS updater archives", { skip: process.platform === "win32" }, () => {
  // No real macOS updater tarball; only an AppImage tarball, which must not satisfy the macOS
  // updater archive requirement. Cover both lower/upper-case variants and both .tar.gz/.tgz.
  const appimageTarballs = ["Formula.appimage.tar.gz", "Formula.AppImage.tgz"];

  for (const tarball of appimageTarballs) {
    const tmp = mkdtempSync(path.join(os.tmpdir(), "formula-checksums-"));
    try {
      writeFixture(tmp, { macUpdaterTarball: tarball });
      const outPath = path.join(tmp, "SHA256SUMS.txt");
      const proc = runGenerateChecksums(tmp, outPath);
      assert.notEqual(proc.status, 0, `expected script to fail when only ${tarball} exists`);
      assert.match(proc.stderr, /\*\.app\.tar\.gz/i);
    } finally {
      rmSync(tmp, { recursive: true, force: true });
    }
  }
});
