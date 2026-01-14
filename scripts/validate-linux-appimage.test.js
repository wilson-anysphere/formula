import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, utimesSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfig = JSON.parse(
  readFileSync(join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"),
);
const expectedVersion = String(tauriConfig?.version ?? "").trim();
const expectedMainBinaryName = String(tauriConfig?.mainBinaryName ?? "").trim() || "formula-desktop";
const expectedFileAssociationMimeTypes = Array.from(
  new Set(
    (tauriConfig?.bundle?.fileAssociations ?? [])
      .flatMap((assoc) => {
        const raw = assoc?.mimeType;
        if (Array.isArray(raw)) return raw;
        if (raw) return [raw];
        return [];
      })
      .map((mt) => String(mt).trim())
      .filter(Boolean),
  ),
);

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

const hasFile = (() => {
  if (!hasBash) return false;
  const probe = spawnSync("file", ["--version"], { stdio: "ignore" });
  return !probe.error && probe.status === 0;
})();

/**
 * Create a fake AppImage executable that supports `--appimage-extract` by creating
 * a synthetic AppDir tree in the current working directory (mimicking AppImageKit).
 *
 * The validator under test executes the AppImage from its own temp dir and expects
 * `./squashfs-root/...` to appear there.
 */
function writeFakeAppImage(
  appImagePath,
  {
    withDesktopFile = true,
    withXlsxMime = true,
    withParquetMime = true,
    withMimeTypeEntry = true,
    withSchemeMime = true,
    withParquetMimeDefinition = true,
    parquetMimeDefinitionContents = "",
    mainBinaryName = expectedMainBinaryName,
    execLine = `${expectedMainBinaryName} %U`,
    withLicense = true,
    withNotice = true,
    appImageVersion = expectedVersion,
    desktopEntryVersion = "",
  } = {},
) {
  let mimeTypes = expectedFileAssociationMimeTypes;
  if (!withXlsxMime) {
    mimeTypes = mimeTypes.filter((mt) => mt !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }
  if (!withParquetMime) {
    mimeTypes = mimeTypes.filter((mt) => mt !== "application/vnd.apache.parquet");
  }
  const desktopMimeBase = `${mimeTypes.join(";")};`;
  const desktopMime = withSchemeMime
    ? `${desktopMimeBase}x-scheme-handler/formula;`
    : desktopMimeBase;

  const desktopBlock = withDesktopFile
    ? [
        "mkdir -p squashfs-root/usr/share/applications",
        "cat > squashfs-root/usr/share/applications/formula.desktop <<'DESKTOP'",
        "[Desktop Entry]",
        "Name=Formula",
        `Exec=${execLine}`,
        ...(appImageVersion ? [`X-AppImage-Version=${appImageVersion}`] : []),
        ...(desktopEntryVersion ? [`Version=${desktopEntryVersion}`] : []),
        ...(withMimeTypeEntry ? [`MimeType=${desktopMime}`] : []),
        "DESKTOP",
      ].join("\n")
    : "mkdir -p squashfs-root/usr/share/applications";

  const defaultParquetMimeDefinitionContents = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
    '  <mime-type type="application/vnd.apache.parquet">',
    "    <comment>Parquet File</comment>",
    '    <glob pattern="*.parquet"/>',
    "  </mime-type>",
    "</mime-info>",
  ].join("\n");
  const parquetMimeContents = parquetMimeDefinitionContents || defaultParquetMimeDefinitionContents;

  const script = `#!/usr/bin/env bash
set -euo pipefail

if [[ "\${1:-}" == "--appimage-extract" ]]; then
  mkdir -p squashfs-root/usr/bin
  mkdir -p squashfs-root/usr/share/doc/${mainBinaryName}
  mkdir -p squashfs-root/usr/share/mime/packages

  cat > squashfs-root/AppRun <<'APPRUN'
#!/usr/bin/env bash
echo "AppRun stub"
APPRUN
  chmod +x squashfs-root/AppRun

  cat > squashfs-root/usr/bin/${mainBinaryName} <<'BIN'
#!/usr/bin/env bash
echo "${mainBinaryName} stub"
BIN
  chmod +x squashfs-root/usr/bin/${mainBinaryName}

  ${withLicense ? 'echo "LICENSE stub" > squashfs-root/usr/share/doc/' + mainBinaryName + '/LICENSE' : ":"}
  ${withNotice ? 'echo "NOTICE stub" > squashfs-root/usr/share/doc/' + mainBinaryName + '/NOTICE' : ":"}
  ${
    withParquetMimeDefinition
      ? `cat > squashfs-root/usr/share/mime/packages/app.formula.desktop.xml <<'MIME'\n${parquetMimeContents}\nMIME`
      : ":"
  }

  ${desktopBlock}
  exit 0
fi

echo "unsupported args: $*" >&2
exit 1
`;

  writeFileSync(appImagePath, script, { encoding: "utf8" });
  chmodSync(appImagePath, 0o755);
}

function runValidator(appImagePath) {
  return runValidatorWithArgs(appImagePath);
}

function runValidatorWithArgs(appImagePath, { args = [], env = {} } = {}) {
  const proc = spawnSync(
    "bash",
    [join(repoRoot, "scripts", "validate-linux-appimage.sh"), "--appimage", appImagePath, ...args],
    {
      cwd: repoRoot,
      encoding: "utf8",
      env: { ...process.env, ...env },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
}

function runValidatorDiscover({ args = [], env = {} } = {}) {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-linux-appimage.sh"), ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env, ...env },
  });
  if (proc.error) throw proc.error;
  return proc;
}

function writeMinimalElf64(path, eMachine) {
  // Minimal ELF64 little endian executable header (enough for `file(1)` to identify arch).
  const buf = Buffer.alloc(64, 0);
  buf[0] = 0x7f;
  buf[1] = 0x45;
  buf[2] = 0x4c;
  buf[3] = 0x46;
  buf[4] = 2; // ELFCLASS64
  buf[5] = 1; // little endian
  buf[6] = 1; // version
  buf.writeUInt16LE(2, 16); // e_type = ET_EXEC
  buf.writeUInt16LE(eMachine, 18); // e_machine
  buf.writeUInt32LE(1, 20); // e_version
  buf.writeUInt16LE(64, 52); // e_ehsize
  buf.writeUInt16LE(56, 54); // e_phentsize
  buf.writeUInt16LE(64, 58); // e_shentsize
  writeFileSync(path, buf);
  chmodSync(path, 0o755);
}

test("validate-linux-appimage --help prints usage and exits 0", { skip: !hasBash }, () => {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-linux-appimage.sh"), "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env },
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /Usage:/);
  assert.doesNotMatch(proc.stderr, /command not found/i);
});

test("validate-linux-appimage accepts a structurally valid AppImage", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-appimage honors FORMULA_APPIMAGE_MAIN_BINARY override", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  const overrideBin = "formula-desktop-alt";
  writeFakeAppImage(appImagePath, {
    mainBinaryName: overrideBin,
    execLine: `${overrideBin} %U`,
    withDesktopFile: true,
    withXlsxMime: true,
    appImageVersion: expectedVersion,
  });

  const proc = runValidatorWithArgs(appImagePath, { env: { FORMULA_APPIMAGE_MAIN_BINARY: overrideBin } });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-appimage accepts when .desktop Exec= wraps the binary in quotes", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, {
    withDesktopFile: true,
    withXlsxMime: true,
    execLine: `\"/usr/bin/${expectedMainBinaryName}\" %U`,
    appImageVersion: expectedVersion,
  });

  const proc = runValidator(appImagePath);
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-appimage auto-discovery validates newest AppImage by default (and --all validates all)", { skip: !hasBash }, () => {
  const tmpTarget = mkdtempSync(join(tmpdir(), "formula-appimage-target-"));
  const bundleDir = join(tmpTarget, "release", "bundle", "appimage");
  mkdirSync(bundleDir, { recursive: true });

  const oldApp = join(bundleDir, "Old.AppImage");
  const newApp = join(bundleDir, "New.AppImage");

  // Old AppImage is intentionally invalid (missing NOTICE).
  writeFakeAppImage(oldApp, {
    withDesktopFile: true,
    withXlsxMime: true,
    withNotice: false,
    appImageVersion: expectedVersion,
  });
  writeFakeAppImage(newApp, { withDesktopFile: true, withXlsxMime: true, appImageVersion: expectedVersion });

  // Make sure the "new" one is selected by mtime.
  utimesSync(oldApp, new Date(1_000_000), new Date(1_000_000));
  utimesSync(newApp, new Date(2_000_000), new Date(2_000_000));

  const procDefault = runValidatorDiscover({ env: { CARGO_TARGET_DIR: tmpTarget } });
  rmSync(tmpTarget, { recursive: true, force: true });
  assert.equal(procDefault.status, 0, procDefault.stderr);

  // Recreate for the --all run (since we rmSync'd the directory above).
  const tmpTarget2 = mkdtempSync(join(tmpdir(), "formula-appimage-target-"));
  const bundleDir2 = join(tmpTarget2, "release", "bundle", "appimage");
  mkdirSync(bundleDir2, { recursive: true });
  const oldApp2 = join(bundleDir2, "Old.AppImage");
  const newApp2 = join(bundleDir2, "New.AppImage");
  writeFakeAppImage(oldApp2, { withDesktopFile: true, withXlsxMime: true, withNotice: false, appImageVersion: expectedVersion });
  writeFakeAppImage(newApp2, { withDesktopFile: true, withXlsxMime: true, appImageVersion: expectedVersion });
  utimesSync(oldApp2, new Date(1_000_000), new Date(1_000_000));
  utimesSync(newApp2, new Date(2_000_000), new Date(2_000_000));

  const procAll = runValidatorDiscover({ args: ["--all"], env: { CARGO_TARGET_DIR: tmpTarget2 } });
  rmSync(tmpTarget2, { recursive: true, force: true });
  assert.notEqual(procAll.status, 0, "expected non-zero exit status");
  assert.match(procAll.stderr, /Missing compliance file/i);
});

test("validate-linux-appimage fails fast on explicit wrong-arch ELF AppImage", { skip: !hasBash || !hasFile }, () => {
  const uname = spawnSync("uname", ["-m"], { encoding: "utf8" });
  if (uname.error || uname.status !== 0) return;
  const arch = (uname.stdout ?? "").trim();

  // Only test on common arch values where we know how to pick an opposite.
  let wrongMachine = null;
  if (arch === "x86_64") {
    wrongMachine = 183; // AArch64
  } else if (arch === "aarch64" || arch === "arm64") {
    wrongMachine = 62; // x86-64
  } else {
    return;
  }

  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-elf-"));
  const appImagePath = join(tmp, "WrongArch.AppImage");
  writeMinimalElf64(appImagePath, wrongMachine);

  const proc = runValidatorWithArgs(appImagePath);
  rmSync(tmp, { recursive: true, force: true });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Wrong AppImage architecture/i);
});

test("validate-linux-appimage --exec-check succeeds on a runnable AppRun", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, appImageVersion: expectedVersion });

  const proc = runValidatorWithArgs(appImagePath, {
    args: ["--exec-check", "--exec-timeout", "2"],
    // Avoid xvfb-run-safe selection and any dependency on Xvfb for this unit test.
    env: { CI: "", DISPLAY: ":99" },
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-appimage fails when no .desktop files exist", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: false, withXlsxMime: true });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /No \.desktop files found/i);
});

test("validate-linux-appimage fails when .desktop lacks xlsx integration", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: false, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /MimeType=.*xlsx|spreadsheet/i);
});

test("validate-linux-appimage fails when .desktop lacks Parquet MIME type", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, {
    withDesktopFile: true,
    withXlsxMime: true,
    withParquetMime: false,
    appImageVersion: expectedVersion,
  });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Parquet support/i);
});

test("validate-linux-appimage fails when .desktop lacks URL scheme handler (x-scheme-handler/formula)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, {
    withDesktopFile: true,
    withXlsxMime: true,
    withSchemeMime: false,
    appImageVersion: expectedVersion,
  });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /x-scheme-handler\/formula/i);
});

test("validate-linux-appimage fails when .desktop lacks a MimeType entry", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withMimeTypeEntry: false, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /MimeType=/i);
});

test("validate-linux-appimage fails when LICENSE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, withLicense: false, withNotice: true, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing compliance file/i);
  assert.match(proc.stderr, /LICENSE/i);
});

test("validate-linux-appimage fails when NOTICE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, withLicense: true, withNotice: false, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Missing compliance file/i);
  assert.match(proc.stderr, /NOTICE/i);
});

test("validate-linux-appimage fails when Parquet shared-mime-info definition is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, {
    withDesktopFile: true,
    withXlsxMime: true,
    withParquetMimeDefinition: false,
    appImageVersion: expectedVersion,
  });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Parquet/i);
  assert.match(proc.stderr, /shared-mime-info/i);
});

test(
  "validate-linux-appimage fails when Parquet shared-mime-info definition is missing expected content",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
    const appImagePath = join(tmp, "Formula.AppImage");
    writeFakeAppImage(appImagePath, {
      withDesktopFile: true,
      withXlsxMime: true,
      withParquetMimeDefinition: true,
      parquetMimeDefinitionContents: [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
        '  <mime-type type="application/vnd.apache.parquet">',
        "    <comment>Parquet File</comment>",
        // Intentionally omit the `*.parquet` glob mapping.
        "  </mime-type>",
        "</mime-info>",
      ].join("\n"),
      appImageVersion: expectedVersion,
    });

    const proc = runValidator(appImagePath);
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /Parquet/i);
    assert.match(proc.stderr, /expected content/i);
  },
);

test("validate-linux-appimage fails when X-AppImage-Version does not match tauri.conf.json", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, appImageVersion: "0.0.0" });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /AppImage version mismatch/i);
});

test("validate-linux-appimage fails when Exec= lacks file placeholder", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, execLine: expectedMainBinaryName });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Exec=.*placeholder|invalid Exec=/i);
});

test("validate-linux-appimage fails when .desktop Exec= does not reference AppRun or the expected binary", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, {
    withDesktopFile: true,
    withXlsxMime: true,
    execLine: "something-else %U",
  });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /target the expected executable/i);
  assert.match(proc.stderr, /AppRun/i);
});

test(
  "validate-linux-appimage accepts Version= when X-AppImage-Version is absent (semver-like)",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
    const appImagePath = join(tmp, "Formula.AppImage");
    writeFakeAppImage(appImagePath, {
      withDesktopFile: true,
      withXlsxMime: true,
      appImageVersion: "",
      desktopEntryVersion: expectedVersion,
    });

    const proc = runValidator(appImagePath);
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "validate-linux-appimage falls back to filename version when no version markers exist",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
    const appImagePath = join(tmp, `Formula-${expectedVersion}.AppImage`);
    writeFakeAppImage(appImagePath, {
      withDesktopFile: true,
      withXlsxMime: true,
      appImageVersion: "",
      desktopEntryVersion: "",
    });

    const proc = runValidator(appImagePath);
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "validate-linux-appimage fails when no version markers exist and filename lacks the version",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
    const appImagePath = join(tmp, "Formula.AppImage");
    writeFakeAppImage(appImagePath, {
      withDesktopFile: true,
      withXlsxMime: true,
      appImageVersion: "",
      desktopEntryVersion: "",
    });

    const proc = runValidator(appImagePath);
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /filename did not contain expected version/i);
  },
);
