import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

const tauriConfig = JSON.parse(
  readFileSync(join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"),
);
const expectedIdentifier = String(tauriConfig?.identifier ?? "").trim();
const expectedVersion = String(tauriConfig?.version ?? "").trim();
const expectedFileExtensions = (() => {
  /** @type {string[]} */
  const out = [];
  const seen = new Set();
  const associations = tauriConfig?.bundle?.fileAssociations ?? [];
  if (!Array.isArray(associations)) return out;

  for (const assoc of associations) {
    if (!assoc || typeof assoc !== "object") continue;
    const raw = assoc.ext;
    const exts = Array.isArray(raw) ? raw : typeof raw === "string" ? [raw] : [];
    for (const ext of exts) {
      if (typeof ext !== "string") continue;
      const normalized = ext.trim().toLowerCase().replace(/^\./, "");
      if (!normalized) continue;
      if (seen.has(normalized)) continue;
      seen.add(normalized);
      out.push(normalized);
    }
  }

  return out;
})();

function writeFakeTool(binDir, name, content) {
  const toolPath = join(binDir, name);
  writeFileSync(toolPath, content, { encoding: "utf8" });
  chmodSync(toolPath, 0o755);
}

function writeFakeMacOsTooling(binDir, { mountPoint, devEntry, lipoArchs, attachLogPath } = {}) {
  writeFakeTool(
    binDir,
    "uname",
    `#!/usr/bin/env bash\nif [[ \"${"$"}{1:-}\" == \"-s\" || \"${"$"}{#}\" -eq 0 ]]; then\n  echo Darwin\n  exit 0\nfi\necho Darwin\n`,
  );

  // macOS `mktemp -t <prefix>` behavior differs from GNU mktemp (which requires
  // a template containing Xs). The validator uses the macOS-style form, so
  // provide a tiny compatibility shim for Linux-based unit tests.
  writeFakeTool(
    binDir,
    "mktemp",
    `#!/usr/bin/env bash\nset -euo pipefail\nis_dir=0\nprefix=\"\"\nwhile [[ \"${"$"}#\" -gt 0 ]]; do\n  case \"${"$"}1\" in\n    -d)\n      is_dir=1\n      shift\n      ;;\n    -t)\n      shift\n      prefix=\"${"$"}{1:-}\"\n      shift || true\n      ;;\n    *)\n      # If a template is provided directly, treat it as the prefix.\n      prefix=\"${"$"}1\"\n      shift\n      ;;\n  esac\ndone\nif [[ -z \"$prefix\" ]]; then\n  prefix=\"tmp\"\nfi\nsuffix=\"${"$"}{RANDOM:-0}.${"$"}{RANDOM:-0}.${"$"}$\"\npath=\"/tmp/${"$"}prefix.${"$"}suffix\"\nif [[ \"$is_dir\" -eq 1 ]]; then\n  mkdir -p \"$path\"\nelse\n  : > \"$path\"\nfi\necho \"$path\"\n`,
  );

  const plist = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n  <dict>\n    <key>system-entities</key>\n    <array>\n      <dict>\n        <key>dev-entry</key>\n        <string>${devEntry}</string>\n        <key>mount-point</key>\n        <string>${mountPoint}</string>\n      </dict>\n    </array>\n  </dict>\n</plist>\n`;

  writeFakeTool(
    binDir,
    "hdiutil",
    `#!/usr/bin/env bash\nset -euo pipefail\ncmd=\"${"$"}{1:-}\"\nshift || true\ncase \"${"$"}cmd\" in\n  attach)\n    ${
      attachLogPath
        ? `log_file=${JSON.stringify(attachLogPath)}\n    dmg=\"\"\n    for arg in \"$@\"; do\n      dmg=\"$arg\"\n    done\n    echo \"$dmg\" >> \"$log_file\"\n\n    `
        : ""
    }# Print plist to stdout; validate-macos-bundle.sh captures it.\n    cat <<'PLIST'\n${plist}PLIST\n    ;;\n  detach)\n    # Accept any detach invocation.\n    exit 0\n    ;;\n  *)\n    echo \"fake hdiutil: unsupported command: ${"$"}cmd\" >&2\n    exit 2\n    ;;\nesac\n`,
  );

  writeFakeTool(
    binDir,
    "lipo",
    `#!/usr/bin/env bash\nset -euo pipefail\nif [[ \"${"$"}{1:-}\" == \"-info\" ]]; then\n  echo \"Architectures in the fat file: ${"$"}{2:-unknown} are: ${(Array.isArray(lipoArchs) && lipoArchs.length > 0 ? lipoArchs : ["x86_64", "arm64"]).join(" ")}\"\n  exit 0\nfi\necho \"fake lipo: unsupported args: ${"$"}*\" >&2\nexit 2\n`,
  );
}

function writeInfoPlist(
  plistPath,
  {
    identifier,
    version,
    urlSchemes = ["formula"],
    fileExtensions = expectedFileExtensions.length > 0 ? expectedFileExtensions : ["xlsx", "xls", "csv"],
  },
) {
  const schemesXml = urlSchemes.map((scheme) => `        <string>${scheme}</string>`).join("\n");
  const extsXml = fileExtensions.map((ext) => `        <string>${ext}</string>`).join("\n");
  const content = `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n<plist version=\"1.0\">\n<dict>\n  <key>CFBundleIdentifier</key>\n  <string>${identifier}</string>\n  <key>CFBundleShortVersionString</key>\n  <string>${version}</string>\n  <key>CFBundleExecutable</key>\n  <string>formula-desktop</string>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n${schemesXml}\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n${extsXml}\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
  writeFileSync(plistPath, content, { encoding: "utf8" });
}

function writeComplianceResources(contentsDir, { includeLicense = true, includeNotice = true } = {}) {
  const resourcesDir = join(contentsDir, "Resources");
  mkdirSync(resourcesDir, { recursive: true });
  if (includeLicense) writeFileSync(join(resourcesDir, "LICENSE"), "stub", { encoding: "utf8" });
  if (includeNotice) writeFileSync(join(resourcesDir, "NOTICE"), "stub", { encoding: "utf8" });
}

function runValidator({ dmgPath, binDir, env = {} }) {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-macos-bundle.sh"), "--dmg", dmgPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      ...env,
      PATH: `${binDir}:${process.env.PATH}`,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

function runValidatorAuto({ binDir, env = {} }) {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-macos-bundle.sh")], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      ...env,
      PATH: `${binDir}:${process.env.PATH}`,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test(
  "validate-macos-bundle validates Info.plist identifier + version metadata",
  { skip: !hasBash },
  () => {
    assert.ok(expectedIdentifier, "tauri.conf.json identifier must be non-empty for this test");
    assert.ok(expectedVersion, "tauri.conf.json version must be non-empty for this test");

    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    // Create a fake app bundle in the "mounted" DMG directory.
    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test("validate-macos-bundle honors FORMULA_TAURI_CONF_PATH (override identifier/version/productName)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });

  const mountPoint = join(tmp, "mnt");
  const devEntry = "/dev/disk99s1";
  mkdirSync(mountPoint, { recursive: true });
  writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

  const overrideIdentifier = "com.example.formula.override";
  const overrideVersion = "0.0.0";
  const overrideProductName = "FormulaOverride";
  const overrideScheme = "formulaoverride";

  const confParent = join(repoRoot, "target");
  mkdirSync(confParent, { recursive: true });
  const confDir = mkdtempSync(join(confParent, "tauri-conf-override-"));
  const confPath = join(confDir, "tauri.conf.json");
  writeFileSync(
    confPath,
    JSON.stringify(
      {
        identifier: overrideIdentifier,
        version: overrideVersion,
        productName: overrideProductName,
        mainBinaryName: "formula-desktop",
        plugins: { "deep-link": { desktop: { schemes: [overrideScheme] } } },
        bundle: { fileAssociations: [] },
      },
      null,
      2,
    ),
    "utf8",
  );

  try {
    const appRoot = join(mountPoint, `${overrideProductName}.app`, "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: overrideIdentifier,
      version: overrideVersion,
      urlSchemes: [overrideScheme],
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({
      dmgPath,
      binDir,
      env: {
        FORMULA_TAURI_CONF_PATH: relative(repoRoot, confPath),
      },
    });
    assert.equal(proc.status, 0, proc.stderr);
  } finally {
    rmSync(confDir, { recursive: true, force: true });
  }
});

test(
  "validate-macos-bundle fails when DMG is empty",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /empty/i);
  },
);

test(
  "validate-macos-bundle artifact discovery prefers universal-apple-darwin DMGs",
  { skip: !hasBash },
  () => {
    assert.ok(expectedIdentifier, "tauri.conf.json identifier must be non-empty for this test");
    assert.ok(expectedVersion, "tauri.conf.json version must be non-empty for this test");

    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    const attachLogPath = join(tmp, "hdiutil-attach.log");
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry, attachLogPath });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    writeComplianceResources(appRoot);

    const cargoTargetDir = join(tmp, "cargo-target");
    const universalDmgPath = join(
      cargoTargetDir,
      "universal-apple-darwin",
      "release",
      "bundle",
      "dmg",
      "Formula.dmg",
    );
    const x86DmgPath = join(
      cargoTargetDir,
      "x86_64-apple-darwin",
      "release",
      "bundle",
      "dmg",
      "Formula.dmg",
    );
    mkdirSync(dirname(universalDmgPath), { recursive: true });
    mkdirSync(dirname(x86DmgPath), { recursive: true });
    writeFileSync(universalDmgPath, "not-a-real-dmg", { encoding: "utf8" });
    writeFileSync(x86DmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidatorAuto({ binDir, env: { CARGO_TARGET_DIR: cargoTargetDir } });
    assert.equal(proc.status, 0, proc.stderr);

    const attached = readFileSync(attachLogPath, "utf8")
      .trim()
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    assert.deepEqual(attached, [universalDmgPath]);
  },
);

test(
  "validate-macos-bundle validates updater tarball artifacts even when --dmg is provided",
  { skip: !hasBash },
  () => {
    assert.ok(expectedIdentifier, "tauri.conf.json identifier must be non-empty for this test");
    assert.ok(expectedVersion, "tauri.conf.json version must be non-empty for this test");

    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const cargoTargetDir = join(tmp, "cargo-target");
    const archivePath = join(
      cargoTargetDir,
      "universal-apple-darwin",
      "release",
      "bundle",
      "macos",
      "Formula.app.tar.gz",
    );
    mkdirSync(dirname(archivePath), { recursive: true });

    const payloadRoot = join(tmp, "payload");
    const payloadAppContents = join(payloadRoot, "Formula.app", "Contents");
    const payloadMacosDir = join(payloadAppContents, "MacOS");
    mkdirSync(payloadMacosDir, { recursive: true });
    writeFileSync(join(payloadMacosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(payloadMacosDir, "formula-desktop"), 0o755);
    writeInfoPlist(join(payloadAppContents, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    writeComplianceResources(payloadAppContents);

    const tarProc = spawnSync("tar", ["-czf", archivePath, "-C", payloadRoot, "Formula.app"], {
      encoding: "utf8",
    });
    assert.equal(tarProc.status, 0, tarProc.stderr);

    const proc = runValidator({ dmgPath, binDir, env: { CARGO_TARGET_DIR: cargoTargetDir } });
    assert.equal(proc.status, 0, proc.stderr);
    assert.match(proc.stdout, /validating updater archive/i);
    assert.match(proc.stdout, /Formula\.app\.tar\.gz/);
  },
);

test(
  "validate-macos-bundle fails when updater tarball is missing compliance resources",
  { skip: !hasBash },
  () => {
    assert.ok(expectedIdentifier, "tauri.conf.json identifier must be non-empty for this test");
    assert.ok(expectedVersion, "tauri.conf.json version must be non-empty for this test");

    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    // DMG-mounted app bundle is valid (compliance resources present) so the failure
    // is attributable to the updater archive validation.
    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const cargoTargetDir = join(tmp, "cargo-target");
    const archivePath = join(
      cargoTargetDir,
      "universal-apple-darwin",
      "release",
      "bundle",
      "macos",
      "Formula.app.tar.gz",
    );
    mkdirSync(dirname(archivePath), { recursive: true });

    const payloadRoot = join(tmp, "payload");
    const payloadAppContents = join(payloadRoot, "Formula.app", "Contents");
    const payloadMacosDir = join(payloadAppContents, "MacOS");
    mkdirSync(payloadMacosDir, { recursive: true });
    writeFileSync(join(payloadMacosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(payloadMacosDir, "formula-desktop"), 0o755);
    writeInfoPlist(join(payloadAppContents, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });
    // Intentionally omit NOTICE from the updater archive payload.
    writeComplianceResources(payloadAppContents, { includeNotice: false });

    const tarProc = spawnSync("tar", ["-czf", archivePath, "-C", payloadRoot, "Formula.app"], {
      encoding: "utf8",
    });
    assert.equal(tarProc.status, 0, tarProc.stderr);

    const proc = runValidator({ dmgPath, binDir, env: { CARGO_TARGET_DIR: cargoTargetDir } });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /compliance/i);
    assert.match(proc.stderr, /notice/i);
  },
);

test(
  "validate-macos-bundle falls back to mainBinaryName when CFBundleExecutable is missing",
  { skip: !hasBash },
  () => {
    assert.ok(expectedIdentifier, "tauri.conf.json identifier must be non-empty for this test");
    assert.ok(expectedVersion, "tauri.conf.json version must be non-empty for this test");

    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    const schemesXml = `        <string>formula</string>`;
    const extsXml = (expectedFileExtensions.length > 0 ? expectedFileExtensions : ["xlsx", "xls", "csv"])
      .map((ext) => `        <string>${ext}</string>`)
      .join("\n");
    const infoPlistContent = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleIdentifier</key>\n  <string>${expectedIdentifier}</string>\n  <key>CFBundleShortVersionString</key>\n  <string>${expectedVersion}</string>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n${schemesXml}\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n${extsXml}\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
    writeFileSync(join(appRoot, "Info.plist"), infoPlistContent, { encoding: "utf8" });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "validate-macos-bundle fails when the expected URL scheme is missing",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
      urlSchemes: ["wrong"],
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /url scheme/i);
  },
);

test(
  "validate-macos-bundle fails when Info.plist declares an invalid URL scheme value like 'formula://'",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
      urlSchemes: ["formula://"],
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /invalid url scheme/i);
  },
);

test(
  "validate-macos-bundle fails when required file association is missing",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
      fileExtensions: ["txt"],
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /file association/i);
    assert.match(proc.stderr, /xlsx/i);
  },
);

test(
  "validate-macos-bundle fails when CFBundleDocumentTypes is missing",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    const infoPlistContent = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleIdentifier</key>\n  <string>${expectedIdentifier}</string>\n  <key>CFBundleShortVersionString</key>\n  <string>${expectedVersion}</string>\n  <key>CFBundleExecutable</key>\n  <string>formula-desktop</string>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
    writeFileSync(join(appRoot, "Info.plist"), infoPlistContent, { encoding: "utf8" });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /CFBundleDocumentTypes/i);
  },
);

test(
  "validate-macos-bundle accepts extensions present via UT*TypeDeclarations",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    if (!expectedFileExtensions.includes("xlsx")) {
      throw new Error(
        `expected tauri.conf.json bundle.fileAssociations to include xlsx; got: ${expectedFileExtensions.join(", ")}`,
      );
    }

    const docExtensions = expectedFileExtensions.filter((ext) => ext !== "xlsx");
    const docExtXml = docExtensions.map((ext) => `        <string>${ext}</string>`).join("\n");

    const infoPlistContent = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleIdentifier</key>\n  <string>${expectedIdentifier}</string>\n  <key>CFBundleShortVersionString</key>\n  <string>${expectedVersion}</string>\n  <key>CFBundleExecutable</key>\n  <string>formula-desktop</string>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeName</key>\n      <string>Formula Documents</string>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n${docExtXml}\n      </array>\n      <key>LSItemContentTypes</key>\n      <array>\n        <string>org.openxmlformats.spreadsheetml.sheet</string>\n      </array>\n      <key>CFBundleTypeRole</key>\n      <string>Editor</string>\n    </dict>\n  </array>\n  <key>UTImportedTypeDeclarations</key>\n  <array>\n    <dict>\n      <key>UTTypeIdentifier</key>\n      <string>org.openxmlformats.spreadsheetml.sheet</string>\n      <key>UTTypeTagSpecification</key>\n      <dict>\n        <key>public.filename-extension</key>\n        <array>\n          <string>xlsx</string>\n        </array>\n      </dict>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
    writeFileSync(join(appRoot, "Info.plist"), infoPlistContent, { encoding: "utf8" });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "validate-macos-bundle fails when an extension appears only in UT*TypeDeclarations without LSItemContentTypes linkage",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    if (!expectedFileExtensions.includes("xlsx")) {
      throw new Error(
        `expected tauri.conf.json bundle.fileAssociations to include xlsx; got: ${expectedFileExtensions.join(", ")}`,
      );
    }

    const docExtensions = expectedFileExtensions.filter((ext) => ext !== "xlsx");
    const docExtXml = docExtensions.map((ext) => `        <string>${ext}</string>`).join("\n");

    // Declare xlsx only in UT*TypeDeclarations but do not link the UTI via LSItemContentTypes, so
    // Launch Services would not treat the app as a handler for .xlsx.
    const infoPlistContent = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0">\n<dict>\n  <key>CFBundleIdentifier</key>\n  <string>${expectedIdentifier}</string>\n  <key>CFBundleShortVersionString</key>\n  <string>${expectedVersion}</string>\n  <key>CFBundleExecutable</key>\n  <string>formula-desktop</string>\n  <key>CFBundleURLTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleURLSchemes</key>\n      <array>\n        <string>formula</string>\n      </array>\n    </dict>\n  </array>\n  <key>CFBundleDocumentTypes</key>\n  <array>\n    <dict>\n      <key>CFBundleTypeName</key>\n      <string>Formula Documents</string>\n      <key>CFBundleTypeExtensions</key>\n      <array>\n${docExtXml}\n      </array>\n      <key>CFBundleTypeRole</key>\n      <string>Editor</string>\n    </dict>\n  </array>\n  <key>UTImportedTypeDeclarations</key>\n  <array>\n    <dict>\n      <key>UTTypeIdentifier</key>\n      <string>org.openxmlformats.spreadsheetml.sheet</string>\n      <key>UTTypeTagSpecification</key>\n      <dict>\n        <key>public.filename-extension</key>\n        <array>\n          <string>xlsx</string>\n        </array>\n      </dict>\n    </dict>\n  </array>\n</dict>\n</plist>\n`;
    writeFileSync(join(appRoot, "Info.plist"), infoPlistContent, { encoding: "utf8" });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /missing extension/i);
    assert.match(proc.stderr, /xlsx/i);
    assert.match(proc.stderr, /LSItemContentTypes/i);
  },
);

test(
  "validate-macos-bundle fails when compliance resources are missing",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    // Intentionally omit LICENSE (required).
    writeComplianceResources(appRoot, { includeLicense: false });

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /compliance/i);
    assert.match(proc.stderr, /license/i);
  },
);

test(
  "validate-macos-bundle fails when the app binary is not universal",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry, lipoArchs: ["x86_64"] });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /arm64/i);
    assert.match(proc.stderr, /slice/i);
  },
);

test(
  "validate-macos-bundle fails on CFBundleIdentifier mismatch",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: "com.example.wrong",
      version: expectedVersion,
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identity metadata mismatch/i);
    assert.match(proc.stderr, /CFBundleIdentifier/i);
  },
);

test(
  "validate-macos-bundle fails on CFBundleShortVersionString mismatch",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: "0.0.0",
    });
    writeComplianceResources(appRoot);

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({ dmgPath, binDir });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identity metadata mismatch/i);
    assert.match(proc.stderr, /CFBundleShortVersionString/i);
  },
);

test(
  "validate-macos-bundle runs codesign + spctl when APPLE_CERTIFICATE is set",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const logPath = join(tmp, "invocations.log");
    writeFakeTool(
      binDir,
      "codesign",
      `#!/usr/bin/env bash\nset -euo pipefail\necho \"codesign $*\" >> \"${logPath}\"\n# validate-macos-bundle.sh probes hardened runtime metadata via:\n#   codesign -dv --verbose=4 <app>\n# Simulate a hardened-runtime signed binary so the validator can pass in unit tests.\ncase \"${"$"}*\" in\n  *\"-dv\"*)\n    echo \"Runtime Version=13.0.0\" >&2\n    echo \"CodeDirectory v=20500 size=0 flags=0x10000(runtime) hashes=0 location=embedded\" >&2\n    ;;\nesac\nexit 0\n`,
    );
    writeFakeTool(
      binDir,
      "spctl",
      `#!/usr/bin/env bash\nset -euo pipefail\necho \"spctl $*\" >> \"${logPath}\"\nexit 0\n`,
    );

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({
      dmgPath,
      binDir,
      env: {
        APPLE_CERTIFICATE: "dummy",
      },
    });
    assert.equal(proc.status, 0, proc.stderr);

    const log = readFileSync(logPath, "utf8");
    assert.match(log, /codesign --verify/, "expected codesign verify invocation");
    assert.match(log, /spctl --assess --type execute/, "expected spctl execute assessment");
  },
);

test(
  "validate-macos-bundle runs stapler validation when notarization env vars are set",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-macos-bundle-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });

    const mountPoint = join(tmp, "mnt");
    const devEntry = "/dev/disk99s1";
    mkdirSync(mountPoint, { recursive: true });
    writeFakeMacOsTooling(binDir, { mountPoint, devEntry });

    const logPath = join(tmp, "invocations.log");
    writeFakeTool(
      binDir,
      "xcrun",
      `#!/usr/bin/env bash\nset -euo pipefail\necho \"xcrun $*\" >> \"${logPath}\"\nexit 0\n`,
    );
    writeFakeTool(
      binDir,
      "spctl",
      `#!/usr/bin/env bash\nset -euo pipefail\necho \"spctl $*\" >> \"${logPath}\"\nexit 0\n`,
    );

    const appRoot = join(mountPoint, "Formula.app", "Contents");
    const macosDir = join(appRoot, "MacOS");
    mkdirSync(macosDir, { recursive: true });
    writeFileSync(join(macosDir, "formula-desktop"), "stub", { encoding: "utf8" });
    chmodSync(join(macosDir, "formula-desktop"), 0o755);
    writeComplianceResources(appRoot);

    writeInfoPlist(join(appRoot, "Info.plist"), {
      identifier: expectedIdentifier,
      version: expectedVersion,
    });

    const dmgPath = join(tmp, "Formula.dmg");
    writeFileSync(dmgPath, "not-a-real-dmg", { encoding: "utf8" });

    const proc = runValidator({
      dmgPath,
      binDir,
      env: {
        APPLE_ID: "user@example.com",
        APPLE_PASSWORD: "app-specific-password",
      },
    });
    assert.equal(proc.status, 0, proc.stderr);

    const log = readFileSync(logPath, "utf8");
    assert.match(log, /xcrun stapler validate .*Formula\.app/, "expected stapler validate app");
    assert.match(log, /xcrun stapler validate .*Formula\.dmg/, "expected stapler validate dmg");
    assert.match(log, /spctl -a -vv --type open/, "expected spctl open assessment");
  },
);
