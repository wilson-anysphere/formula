import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "check-desktop-url-scheme.mjs");

const defaultIdentifier = "app.formula.desktop";
const requiredFileExtensions = [
  "xlsx",
  "xls",
  "xlt",
  "xla",
  "xlsm",
  "xltx",
  "xltm",
  "xlam",
  "xlsb",
  "csv",
  "parquet",
];

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function isParquetConfigured(config) {
  const assocs = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  return assocs.some((assoc) => {
    const raw = assoc?.ext;
    const exts = Array.isArray(raw) ? raw : typeof raw === "string" ? [raw] : [];
    return exts.some((e) => String(e).trim().toLowerCase().replace(/^\./, "") === "parquet");
  });
}

function runWithConfigAndPlist(config, plistContents) {
  const tmpRoot = path.join(repoRoot, ".tmp");
  mkdirSync(tmpRoot, { recursive: true });
  const dir = mkdtempSync(path.join(tmpRoot, "check-desktop-url-scheme-"));

  const configPath = path.join(dir, "tauri.conf.json");
  writeFileSync(configPath, `${JSON.stringify(config)}\n`, "utf8");

  const plistPath = path.join(dir, "Info.plist");
  writeFileSync(plistPath, plistContents, "utf8");

  // The preflight script validates that when Parquet is configured we ship a shared-mime-info
  // definition under `mime/<identifier>.xml` (relative to tauri.conf.json). Create a minimal
  // definition file in the synthetic config directory so the preflight can read it.
  const parquetConfigured = isParquetConfigured(config);
  const skipMimeDefinition = Boolean(config?.__testSkipParquetMimeDefinition);
  const overrideMimeXml = typeof config?.__testParquetMimeXml === "string" ? config.__testParquetMimeXml : "";
  if (parquetConfigured && !skipMimeDefinition) {
    const identifier =
      typeof config?.identifier === "string" && config.identifier.trim() ? config.identifier.trim() : defaultIdentifier;
    const mimeDir = path.join(dir, "mime");
    mkdirSync(mimeDir, { recursive: true });
    writeFileSync(
      path.join(mimeDir, `${identifier}.xml`),
      overrideMimeXml ||
        [
          '<?xml version="1.0" encoding="UTF-8"?>',
          '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
          '  <mime-type type="application/vnd.apache.parquet">',
          '    <glob pattern="*.parquet" />',
          "  </mime-type>",
          "</mime-info>",
          "",
        ].join("\n"),
      "utf8",
    );
  }

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

function basePlistWithFormulaScheme(fileExtensions = requiredFileExtensions) {
  const extsXml = fileExtensions.map((ext) => `        <string>${ext}</string>`).join("\n");
  return `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n${extsXml}\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
}

const defaultFileAssociations = [
  { ext: ["xlsx"], mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { ext: ["xls"], mimeType: "application/vnd.ms-excel" },
  { ext: ["xlt"], mimeType: "application/vnd.ms-excel" },
  { ext: ["xla"], mimeType: "application/vnd.ms-excel" },
  { ext: ["xlsm"], mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12" },
  { ext: ["xltx"], mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template" },
  { ext: ["xltm"], mimeType: "application/vnd.ms-excel.template.macroEnabled.12" },
  { ext: ["xlam"], mimeType: "application/vnd.ms-excel.addin.macroEnabled.12" },
  { ext: ["xlsb"], mimeType: "application/vnd.ms-excel.sheet.binary.macroEnabled.12" },
  { ext: ["csv"], mimeType: "text/csv" },
  { ext: ["parquet"], mimeType: "application/vnd.apache.parquet" },
];

const parquetMimeDest = `usr/share/mime/packages/${defaultIdentifier}.xml`;
const parquetMimeSrc = `mime/${defaultIdentifier}.xml`;

const defaultLinuxBundle = {
  deb: {
    depends: ["shared-mime-info"],
    files: { [parquetMimeDest]: parquetMimeSrc },
  },
  rpm: {
    depends: ["shared-mime-info"],
    files: { [parquetMimeDest]: parquetMimeSrc },
  },
  appimage: {
    files: { [parquetMimeDest]: parquetMimeSrc },
  },
};

function baseConfig({ fileAssociations } = {}) {
  return {
    identifier: defaultIdentifier,
    plugins: {
      "deep-link": {
        desktop: { schemes: ["formula"] },
      },
    },
    bundle: {
      fileAssociations: fileAssociations ?? cloneJson(defaultFileAssociations),
      linux: cloneJson(defaultLinuxBundle),
    },
  };
}

test("passes when bundle.fileAssociations includes required extensions and all entries have mimeType", () => {
  const config = baseConfig();
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("passes when deep-link schemes is configured as a string", () => {
  const config = baseConfig();
  config.plugins["deep-link"].desktop.schemes = "formula";
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("passes when deep-link schemes includes 'formula://' (normalized)", () => {
  const config = baseConfig();
  config.plugins["deep-link"].desktop.schemes = ["formula://"];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("fails when macOS Info.plist declares an invalid URL scheme value like 'formula://'", () => {
  const config = baseConfig();
  const plist = basePlistWithFormulaScheme().replace("<string>formula</string>", "<string>formula://</string>");
  const proc = runWithConfigAndPlist(config, plist);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Invalid macOS URL scheme registration/i);
  assert.match(proc.stderr, /formula:\/\//i);
});

test("passes when deep-link desktop config is an array of protocol objects", () => {
  const config = baseConfig();
  config.plugins["deep-link"].desktop = [{ schemes: ["formula"] }];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /preflight passed/i);
});

test("fails when Parquet association is configured but bundle.linux is missing", () => {
  const config = baseConfig();
  delete config.bundle.linux;
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Parquet file association configured/i);
  assert.match(proc.stderr, /bundle\.linux/i);
});

test("fails when Parquet association is configured but identifier is missing", () => {
  const config = baseConfig();
  delete config.identifier;
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /identifier is missing/i);
});

test("fails when Parquet association is configured but identifier contains a path separator", () => {
  const config = baseConfig();
  // Backslash is a path separator on Windows and should be rejected even when tests run on Linux/macOS.
  config.identifier = "app\\formula.desktop";
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /identifier is not a valid filename/i);
});

test("fails when Parquet shared-mime-info file mapping does not match identifier (Linux bundle files)", () => {
  const config = baseConfig();
  config.bundle.linux.deb.files[parquetMimeDest] = "mime/wrong.xml";
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /mapping mismatch/i);
});

test("fails when Parquet shared-mime-info mapping is incorrect for RPM bundle files", () => {
  const config = baseConfig();
  config.bundle.linux.rpm.files[parquetMimeDest] = "mime/wrong.xml";
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /mapping mismatch/i);
});

test("fails when Parquet is configured but shared-mime-info is not declared as a DEB dependency", () => {
  const config = baseConfig();
  config.bundle.linux.deb.depends = ["libgtk-3-0"];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info is not declared as a DEB dependency/i);
});

test("fails when Parquet is configured but shared-mime-info is not declared as an RPM dependency", () => {
  const config = baseConfig();
  config.bundle.linux.rpm.depends = ["gtk3"];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info is not declared as an RPM dependency/i);
});

test("fails when Parquet is configured but the shared-mime-info definition file is missing", () => {
  const config = baseConfig();
  config.__testSkipParquetMimeDefinition = true;
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info definition file is missing/i);
});

test("fails when Parquet is configured but shared-mime-info definition file lacks expected content", () => {
  const config = baseConfig();
  config.__testParquetMimeXml = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
    '  <mime-type type="application/vnd.apache.parquet">',
    // Intentionally omit the *.parquet glob mapping.
    "  </mime-type>",
    "</mime-info>",
    "",
  ].join("\n");
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing expected content/i);
  assert.match(proc.stderr, /glob pattern/i);
});

test("fails when bundle.fileAssociations is present but missing required extensions", () => {
  const config = baseConfig({
    fileAssociations: [
      { ext: ["csv"], mimeType: "text/csv" },
      { ext: ["parquet"], mimeType: "application/vnd.apache.parquet" },
    ],
  });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /required desktop file associations/i);
  assert.match(proc.stderr, /xlsx/i);
});

test("fails when macOS Info.plist does not declare the formula:// URL scheme", () => {
  const config = baseConfig();
  const plist = basePlistWithFormulaScheme().replace("<string>formula</string>", "<string>wrong</string>");
  const proc = runWithConfigAndPlist(config, plist);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing macOS URL scheme registration/i);
  assert.match(proc.stderr, /Info\.plist/i);
});

test("passes when macOS Info.plist declares multiple URL schemes and includes formula (not necessarily first)", () => {
  const config = baseConfig();
  const extsXml = requiredFileExtensions.map((ext) => `        <string>${ext}</string>`).join("\n");
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleURLTypes</key>
  <array>
    <dict>
      <key>CFBundleURLSchemes</key>
      <array>
        <string>wrong</string>
      </array>
    </dict>
    <dict>
      <key>CFBundleURLSchemes</key>
      <array>
        <string>formula</string>
      </array>
    </dict>
  </array>
  <key>CFBundleDocumentTypes</key>
  <array>
    <dict>
      <key>CFBundleTypeExtensions</key>
      <array>
${extsXml}
      </array>
    </dict>
  </array>
</dict>
</plist>
`;
  const proc = runWithConfigAndPlist(config, plist);
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when tauri.conf.json deep-link schemes includes an extra scheme not declared in macOS Info.plist", () => {
  const config = baseConfig();
  config.plugins["deep-link"].desktop.schemes = ["formula", "formula-extra"];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing macOS URL scheme registration/i);
  assert.match(proc.stderr, /formula-extra/i);
});

test("fails when tauri.conf.json deep-link schemes do not include formula", () => {
  const config = baseConfig();
  config.plugins["deep-link"].desktop.schemes = ["wrong"];
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing desktop deep-link scheme configuration/i);
  assert.match(proc.stderr, /plugins\["deep-link"\]/i);
});

test("fails when .xlsx association is missing a mimeType entry", () => {
  const fileAssociations = cloneJson(defaultFileAssociations);
  // Remove mimeType for the required .xlsx association.
  fileAssociations[0] = { ext: ["xlsx"] };
  const config = baseConfig({ fileAssociations });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing Linux mimeType fields/i);
  assert.match(proc.stderr, /xlsx/i);
});

test("fails when bundle.fileAssociations contains duplicate extensions", () => {
  const fileAssociations = cloneJson(defaultFileAssociations);
  // Duplicate the .xlsx entry (even with the same mimeType); the preflight should reject it.
  fileAssociations.push({
    ext: ["xlsx"],
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const config = baseConfig({ fileAssociations });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Duplicate file association extensions/i);
  assert.match(proc.stderr, /xlsx/i);
});

test("fails when Parquet association uses an unexpected mimeType", () => {
  const fileAssociations = cloneJson(defaultFileAssociations);
  const parquet = fileAssociations.find((assoc) => Array.isArray(assoc?.ext) && assoc.ext.includes("parquet"));
  if (!parquet) throw new Error("expected defaultFileAssociations to include parquet");
  parquet.mimeType = "application/x-parquet";
  const config = baseConfig({ fileAssociations });
  const proc = runWithConfigAndPlist(config, basePlistWithFormulaScheme());
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /mimeType mismatch/i);
  assert.match(proc.stderr, /parquet/i);
  assert.match(proc.stderr, /application\/vnd\.apache\.parquet/i);
});

test("fails when macOS Info.plist is missing CFBundleDocumentTypes", () => {
  const config = baseConfig();
  const plist = basePlistWithFormulaScheme().replace(/<key>CFBundleDocumentTypes[\s\S]*$/i, "</dict>\n</plist>\n");
  const proc = runWithConfigAndPlist(config, plist);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing macOS file association registration/i);
  assert.match(proc.stderr, /CFBundleDocumentTypes/i);
});

test("fails when macOS Info.plist CFBundleDocumentTypes does not include xlsx", () => {
  const config = baseConfig();
  const plist = basePlistWithFormulaScheme().replace("<string>xlsx</string>", "<string>txt</string>");
  const proc = runWithConfigAndPlist(config, plist);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing macOS file association registration/i);
  assert.match(proc.stderr, /xlsx/i);
});

test("fails when xlsx appears only in UT*TypeDeclarations (not CFBundleDocumentTypes)", () => {
  const config = baseConfig();
  const docTypeExts = requiredFileExtensions.filter((ext) => ext !== "xlsx");
  const extsXml = docTypeExts.map((ext) => `        <string>${ext}</string>`).join("\n");
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleURLTypes</key>
  <array>
    <dict>
      <key>CFBundleURLSchemes</key>
      <array>
        <string>formula</string>
      </array>
    </dict>
  </array>
  <key>CFBundleDocumentTypes</key>
  <array>
    <dict>
      <key>CFBundleTypeExtensions</key>
      <array>
 ${extsXml}
      </array>
    </dict>
  </array>
  <key>UTImportedTypeDeclarations</key>
  <array>
    <dict>
      <key>UTTypeIdentifier</key>
      <string>org.openxmlformats.spreadsheetml.sheet</string>
      <key>UTTypeTagSpecification</key>
      <dict>
        <key>public.filename-extension</key>
        <array>
          <string>xlsx</string>
        </array>
      </dict>
    </dict>
  </array>
</dict>
</plist>
`;

  const proc = runWithConfigAndPlist(config, plist);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing macOS file association registration/i);
});

test("passes when xlsx is registered via LSItemContentTypes + UT*TypeDeclarations (macOS Info.plist)", () => {
  const config = baseConfig();
  const docTypeExts = requiredFileExtensions.filter((ext) => ext !== "xlsx");
  const extsXml = docTypeExts.map((ext) => `        <string>${ext}</string>`).join("\n");
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleURLTypes</key>
  <array>
    <dict>
      <key>CFBundleURLSchemes</key>
      <array>
        <string>formula</string>
      </array>
    </dict>
  </array>
  <key>CFBundleDocumentTypes</key>
  <array>
    <dict>
      <key>CFBundleTypeExtensions</key>
      <array>
${extsXml}
      </array>
      <key>LSItemContentTypes</key>
      <array>
        <string>org.openxmlformats.spreadsheetml.sheet</string>
      </array>
    </dict>
  </array>
  <key>UTImportedTypeDeclarations</key>
  <array>
    <dict>
      <key>UTTypeIdentifier</key>
      <string>org.openxmlformats.spreadsheetml.sheet</string>
      <key>UTTypeTagSpecification</key>
      <dict>
        <key>public.filename-extension</key>
        <array>
          <string>xlsx</string>
        </array>
      </dict>
    </dict>
  </array>
</dict>
</plist>
`;

  const proc = runWithConfigAndPlist(config, plist);
  assert.equal(proc.status, 0, proc.stderr);
});
