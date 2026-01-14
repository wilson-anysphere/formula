import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-desktop-url-scheme.mjs");

function runWithConfigAndPlist(config, plistContents) {
  const tmpRoot = path.join(repoRoot, ".tmp");
  mkdirSync(tmpRoot, { recursive: true });
  const dir = mkdtempSync(path.join(tmpRoot, "check-desktop-url-scheme-"));

  const configPath = path.join(dir, "tauri.conf.json");
  writeFileSync(configPath, `${JSON.stringify(config)}\n`, "utf8");

  const plistPath = path.join(dir, "Info.plist");
  writeFileSync(plistPath, plistContents, "utf8");

  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: configPath,
      FORMULA_INFO_PLIST_PATH: plistPath,
    },
  });
  if (proc.error) throw proc.error;

  rmSync(dir, { recursive: true, force: true });
  return proc;
}

function basePlistWithFormulaScheme() {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
}

function baseConfig({ fileAssociations }) {
  return {
    plugins: {
      "deep-link": {
        desktop: { schemes: ["formula"] },
      },
    },
    bundle: {
      fileAssociations,
    },
  };
}

test("passes when bundle.fileAssociations includes .xlsx and all entries have mimeType", () => {
  const config = baseConfig({
    fileAssociations: [
      { ext: ["xlsx"], mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      { ext: ["csv"], mimeType: "text/csv" },
    ],
  });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("fails when bundle.fileAssociations is present but does not include .xlsx", () => {
  const config = baseConfig({
    fileAssociations: [{ ext: ["csv"], mimeType: "text/csv" }],
  });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /file association configuration/i);
  assert.match(proc.stderr, /xlsx/i);
});

