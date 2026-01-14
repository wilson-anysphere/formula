import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "verify_macos_bundle_associations.py");

const hasPython3 = (() => {
  const probe = spawnSync("python3", ["--version"], { stdio: "ignore" });
  return !probe.error && probe.status === 0;
})();

function writeConfig(dir, { includeDeepLink = true } = {}) {
  const configPath = path.join(dir, "tauri.conf.json");
  const conf = {
    bundle: {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      ],
    },
    ...(includeDeepLink
      ? {
          plugins: {
            "deep-link": {
              desktop: {
                schemes: ["formula"],
              },
            },
          },
        }
      : {}),
  };
  writeFileSync(configPath, JSON.stringify(conf), "utf8");
  return configPath;
}

function writeInfoPlist(dir, { includeXlsxDocumentType = true, includeUrlScheme = true, xlsxInUtiOnly = false } = {}) {
  const plistPath = path.join(dir, "Info.plist");

  const urlBlock = includeUrlScheme
    ? `  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n`
    : "";

  const docTypeExt = includeXlsxDocumentType ? "xlsx" : "txt";
  const docBlock = `  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n        <string>${docTypeExt}</string>\n      </array>\n    </dict>\n  </array>\n`;

  const utiBlock = xlsxInUtiOnly
    ? `  <key>UTImportedTypeDeclarations</key>\n  <array>\n    <dict>\n      <key>UTTypeTagSpecification</key>\n      <dict>\n        <key>public.filename-extension</key>\n        <array>\n          <string>xlsx</string>\n        </array>\n      </dict>\n    </dict>\n  </array>\n`
    : "";

  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n${urlBlock}${docBlock}${utiBlock}</dict>\n</plist>\n`;
  writeFileSync(plistPath, content, "utf8");
  return plistPath;
}

function runValidator({ configPath, infoPlistPath }) {
  const proc = spawnSync(
    "python3",
    [scriptPath, "--tauri-config", configPath, "--info-plist", infoPlistPath],
    { cwd: repoRoot, encoding: "utf8" },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test("verify_macos_bundle_associations passes when Info.plist declares xlsx document types and formula:// scheme", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
  mkdirSync(tmp, { recursive: true });
  const configPath = writeConfig(tmp);
  const infoPlistPath = writeInfoPlist(tmp, { includeXlsxDocumentType: true, includeUrlScheme: true });

  const proc = runValidator({ configPath, infoPlistPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify_macos_bundle_associations fails when CFBundleDocumentTypes does not include xlsx", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
  const configPath = writeConfig(tmp);
  const infoPlistPath = writeInfoPlist(tmp, { includeXlsxDocumentType: false, includeUrlScheme: true });

  const proc = runValidator({ configPath, infoPlistPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing extensions/i);
  assert.match(proc.stderr, /xlsx/i);
});

test("verify_macos_bundle_associations fails when xlsx appears only in UT*TypeDeclarations", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
  const configPath = writeConfig(tmp);
  const infoPlistPath = writeInfoPlist(tmp, {
    includeXlsxDocumentType: false,
    includeUrlScheme: true,
    xlsxInUtiOnly: true,
  });

  const proc = runValidator({ configPath, infoPlistPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /CFBundleDocumentTypes/i);
});

