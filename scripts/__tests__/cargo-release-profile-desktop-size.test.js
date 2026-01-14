import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const cargoTomlPath = path.join(repoRoot, "Cargo.toml");
const desktopCargoTomlPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "Cargo.toml");

async function readCargoToml() {
  // Strip comments so inline `# ...` annotations can't break the guardrail and commented-out
  // settings can't satisfy it.
  return stripHashComments(await readFile(cargoTomlPath, "utf8"));
}

async function readDesktopCargoToml() {
  return stripHashComments(await readFile(desktopCargoTomlPath, "utf8"));
}

/**
 * Best-effort parser for the desktop Tauri crate package name.
 *
 * @param {string} text
 */
function parseDesktopPackageName(text) {
  const lines = text.split(/\r?\n/);
  let inPackage = false;
  for (const raw of lines) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    if (line.startsWith("[") && line.endsWith("]")) {
      inPackage = line === "[package]";
      continue;
    }
    if (!inPackage) continue;
    const m = line.match(/^name\s*=\s*"([^"]+)"/);
    if (m?.[1]) return m[1];
  }
  return null;
}

/**
 * Extracts a TOML section block (e.g. `[profile.release]`) by scanning forward until the
 * next top-level section header.
 *
 * @param {string[]} lines
 * @param {number} startIdx
 */
function tomlSectionBlock(lines, startIdx) {
  let endIdx = startIdx + 1;
  for (; endIdx < lines.length; endIdx += 1) {
    const line = lines[endIdx] ?? "";
    if (/^\s*\[.*\]\s*$/.test(line)) break;
  }
  return lines.slice(startIdx, endIdx).join("\n");
}

test("root Cargo.toml release profile keeps size-focused defaults for shipped desktop artifacts", async () => {
  const text = await readCargoToml();
  const lines = text.split(/\r?\n/);
  const desktopToml = await readDesktopCargoToml();
  const desktopPackageName = parseDesktopPackageName(desktopToml);
  assert.ok(
    desktopPackageName,
    `Expected to parse desktop package name from ${path.relative(repoRoot, desktopCargoTomlPath)}`,
  );

  const releaseHeader = "[profile.release]";
  const releaseIdx = lines.findIndex((line) => line.trim() === releaseHeader);
  assert.ok(releaseIdx >= 0, `Expected ${path.relative(repoRoot, cargoTomlPath)} to include ${releaseHeader}`);
  const releaseBlock = tomlSectionBlock(lines, releaseIdx);

  assert.match(
    releaseBlock,
    /^\s*strip\s*=\s*"symbols"\s*$/m,
    `Expected ${releaseHeader} to set strip = "symbols"`,
  );
  assert.match(
    releaseBlock,
    /^\s*lto\s*=\s*"thin"\s*$/m,
    `Expected ${releaseHeader} to set lto = "thin"`,
  );
  assert.match(
    releaseBlock,
    /^\s*codegen-units\s*=\s*1\s*$/m,
    `Expected ${releaseHeader} to set codegen-units = 1`,
  );

  const desktopHeader = `[profile.release.package.${desktopPackageName}]`;
  const desktopIdx = lines.findIndex((line) => line.trim() === desktopHeader);
  assert.ok(
    desktopIdx >= 0,
    `Expected ${path.relative(repoRoot, cargoTomlPath)} to include ${desktopHeader} (matching desktop package ${desktopPackageName}).`,
  );
  const desktopBlock = tomlSectionBlock(lines, desktopIdx);
  assert.match(
    desktopBlock,
    /^\s*opt-level\s*=\s*"z"\s*$/m,
    `Expected ${desktopHeader} to set opt-level = "z"`,
  );
});
