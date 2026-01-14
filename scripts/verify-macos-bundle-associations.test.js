import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
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

function writeInfoPlist(
  dir,
  {
    includeXlsxDocumentType = true,
    includeUrlScheme = true,
    xlsxInUtiOnly = false,
    includeLsItemContentTypes = false,
  } = {},
) {
  const plistPath = path.join(dir, "Info.plist");

  const urlBlock = includeUrlScheme
    ? `  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n`
    : "";

  const docTypeExt = includeXlsxDocumentType ? "xlsx" : "txt";
  const lsItemBlock = includeLsItemContentTypes
    ? `      <key>LSItemContentTypes</key>\n      <array>\n        <string>org.openxmlformats.spreadsheetml.sheet</string>\n      </array>\n`
    : "";
  const docBlock = `  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n        <string>${docTypeExt}</string>\n      </array>\n${lsItemBlock}    </dict>\n  </array>\n`;

  const utiBlock = xlsxInUtiOnly
    ? `  <key>UTImportedTypeDeclarations</key>\n  <array>\n    <dict>\n      <key>UTTypeIdentifier</key>\n      <string>org.openxmlformats.spreadsheetml.sheet</string>\n      <key>UTTypeTagSpecification</key>\n      <dict>\n        <key>public.filename-extension</key>\n        <array>\n          <string>xlsx</string>\n        </array>\n      </dict>\n    </dict>\n  </array>\n`
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
  assert.match(proc.stderr, /UT\*TypeDeclarations/i);
});

test("verify_macos_bundle_associations passes when xlsx is registered via LSItemContentTypes + UT*TypeDeclarations", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
  const configPath = writeConfig(tmp);
  const infoPlistPath = writeInfoPlist(tmp, {
    includeXlsxDocumentType: false,
    includeUrlScheme: true,
    xlsxInUtiOnly: true,
    includeLsItemContentTypes: true,
  });

  const proc = runValidator({ configPath, infoPlistPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify_macos_bundle_associations supports tauri.conf.json fileAssociations ext as a string", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
  const configPath = path.join(tmp, "tauri.conf.json");
  writeFileSync(
    configPath,
    JSON.stringify({
      bundle: {
        fileAssociations: [
          {
            ext: "xlsx",
            mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
        ],
      },
      plugins: {
        "deep-link": {
          desktop: {
            schemes: ["formula"],
          },
        },
      },
    }),
    "utf8",
  );
  const infoPlistPath = writeInfoPlist(tmp, { includeXlsxDocumentType: true, includeUrlScheme: true });

  const proc = runValidator({ configPath, infoPlistPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test(
  "verify_macos_bundle_associations validates multiple deep-link schemes when desktop config is an array",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify({
        bundle: {
          fileAssociations: [
            {
              ext: ["xlsx"],
              mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
          ],
        },
        plugins: {
          "deep-link": {
            desktop: [
              {
                schemes: ["formula", "formula-extra"],
              },
            ],
          },
        },
      }),
      "utf8",
    );

    const infoPlistPath = path.join(tmp, "Info.plist");
    writeFileSync(
      infoPlistPath,
      `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n        <string>formula-extra</string>\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n        <string>xlsx</string>\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`,
      "utf8",
    );

    const proc = runValidator({ configPath, infoPlistPath });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "verify_macos_bundle_associations normalizes deep-link schemes like formula:// from config",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify({
        bundle: {
          fileAssociations: [
            {
              ext: ["xlsx"],
              mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
          ],
        },
        plugins: {
          "deep-link": {
            desktop: {
              schemes: ["formula://", "formula-extra:"],
            },
          },
        },
      }),
      "utf8",
    );

    const infoPlistPath = path.join(tmp, "Info.plist");
    writeFileSync(
      infoPlistPath,
      `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n        <string>formula-extra</string>\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n        <string>xlsx</string>\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`,
      "utf8",
    );

    const proc = runValidator({ configPath, infoPlistPath });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "verify_macos_bundle_associations fails when Info.plist declares an invalid scheme value like formula://",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
    const configPath = writeConfig(tmp);
    const infoPlistPath = writeInfoPlist(tmp, { includeXlsxDocumentType: true, includeUrlScheme: true });
    const raw = readFileSync(infoPlistPath, "utf8").replace("<string>formula</string>", "<string>formula://</string>");
    writeFileSync(infoPlistPath, raw, "utf8");

    const proc = runValidator({ configPath, infoPlistPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /invalid/i);
    assert.match(proc.stderr, /CFBundleURLSchemes/i);
    assert.match(proc.stderr, /formula:\/\//i);
  },
);

test(
  "verify_macos_bundle_associations fails when a configured deep-link scheme is missing from Info.plist",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-macos-assoc-test-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify({
        bundle: {
          fileAssociations: [
            {
              ext: ["xlsx"],
              mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
          ],
        },
        plugins: {
          "deep-link": {
            desktop: [
              {
                schemes: ["formula", "formula-extra"],
              },
            ],
          },
        },
      }),
      "utf8",
    );

    const infoPlistPath = writeInfoPlist(tmp, { includeXlsxDocumentType: true, includeUrlScheme: true });
    const proc = runValidator({ configPath, infoPlistPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /Missing URL schemes/i);
    assert.match(proc.stderr, /formula-extra/i);
  },
);
