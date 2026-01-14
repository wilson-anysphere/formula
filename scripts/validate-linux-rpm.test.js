import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const tauriConf = JSON.parse(readFileSync(join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"));
const expectedVersion = String(tauriConf?.version ?? "").trim();
const expectedMainBinaryName = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";
const expectedRpmName = expectedMainBinaryName;

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

function writeFakeRpmTool(binDir) {
  const rpmScript = `#!/usr/bin/env bash
 set -euo pipefail

  mode="\${FAKE_RPM_MODE:-ok}"
 fake_version="\${FAKE_RPM_VERSION:-0.0.0}"
 fake_name="\${FAKE_RPM_NAME:-formula-desktop}"

 cmd="\${1:-}"
 if [[ "$cmd" == "-qpR" ]]; then
   rpm_path="\${2:-}"
   if [[ -z "\${FAKE_RPM_REQUIRES_FILE:-}" ]]; then
     echo "fake rpm: missing FAKE_RPM_REQUIRES_FILE" >&2
     exit 2
   fi
   cat "$FAKE_RPM_REQUIRES_FILE"
   exit 0
 fi

 if [[ "$cmd" != "-qp" ]]; then
   echo "fake rpm: unexpected args: $*" >&2
   exit 2
 fi

  query="\${2:-}"

 
 case "$query" in
   --info)
     rpm_path="\${3:-}"
     if [[ "$mode" == "fail-info" ]]; then
       echo "fake rpm: failing --info for $rpm_path" >&2
       exit 1
     fi
     echo "Name        : $fake_name"
     echo "Version     : $fake_version"
     exit 0
     ;;
   --list)
     rpm_path="\${3:-}"
     if [[ "$mode" == "fail-list" ]]; then
       echo "fake rpm: failing --list for $rpm_path" >&2
       exit 1
    fi
    if [[ -z "\${FAKE_RPM_LIST_FILE:-}" ]]; then
      echo "fake rpm: missing FAKE_RPM_LIST_FILE" >&2
      exit 2
    fi
    cat "$FAKE_RPM_LIST_FILE"
    exit 0
    ;;
  --queryformat)
    fmt="\${3:-}"
    rpm_path="\${4:-}"
    if [[ "$mode" == "fail-queryformat" ]]; then
      echo "fake rpm: failing --queryformat for $rpm_path" >&2
      exit 1
    fi
    if [[ "$fmt" == *"%{VERSION}"* ]]; then
      echo "$fake_version"
      exit 0
    fi
    if [[ "$fmt" == *"%{NAME}"* ]]; then
      echo "$fake_name"
      exit 0
    fi
    echo "fake rpm: unsupported queryformat: $fmt" >&2
    exit 2
    ;;
  *)
    echo "fake rpm: unsupported query: $query" >&2
    exit 2
    ;;
esac
`;

  const rpmPath = join(binDir, "rpm");
  writeFileSync(rpmPath, rpmScript, { encoding: "utf8" });
  chmodSync(rpmPath, 0o755);
}

function writeDefaultRequiresFile(tmpDir) {
  const requiresPath = join(tmpDir, "rpm-requires.txt");
  writeFileSync(
    requiresPath,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );
  return requiresPath;
}

function writeFakeRpmExtractTools(
  binDir,
  {
    withMimeType = true,
    mimeTypeLine = "MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;",
    execLine = `Exec=${expectedMainBinaryName} %U`,
  } = {},
) {
  const rpm2cpioScript = `#!/usr/bin/env bash
set -euo pipefail
# The validator only uses rpm2cpio as part of a pipe into cpio; the test fakes
# extraction by implementing a fake cpio that writes the desired files.
exit 0
`;
  const rpm2cpioPath = join(binDir, "rpm2cpio");
  writeFileSync(rpm2cpioPath, rpm2cpioScript, { encoding: "utf8" });
  chmodSync(rpm2cpioPath, 0o755);

  let effectiveMimeTypeLine = mimeTypeLine;
  if (withMimeType && !effectiveMimeTypeLine.toLowerCase().includes("x-scheme-handler/")) {
    if (!effectiveMimeTypeLine.trim().endsWith(";")) {
      effectiveMimeTypeLine = `${effectiveMimeTypeLine};`;
    }
    effectiveMimeTypeLine = `${effectiveMimeTypeLine}x-scheme-handler/formula;`;
  }

  const desktopLines = [
    "[Desktop Entry]",
    "Type=Application",
    "Name=Formula",
    execLine,
    ...(withMimeType ? [effectiveMimeTypeLine] : []),
  ];

  const cpioScript = `#!/usr/bin/env bash
set -euo pipefail
# Drain stdin so pipes don't break unexpectedly.
cat >/dev/null || true

mkdir -p usr/share/applications
cat > usr/share/applications/formula.desktop <<'DESKTOP'
${desktopLines.join("\n")}
DESKTOP
exit 0
`;
  const cpioPath = join(binDir, "cpio");
  writeFileSync(cpioPath, cpioScript, { encoding: "utf8" });
  chmodSync(cpioPath, 0o755);
}

function runValidator({
  cwd,
  rpmArg,
  fakeListFile,
  fakeRequiresFile,
  fakeMode,
  fakeVersion,
  fakeName,
  rpmNameOverride,
}) {
  const proc = spawnSync(
    "bash",
    [join(repoRoot, "scripts", "validate-linux-rpm.sh"), "--no-container", "--rpm", rpmArg],
    {
      cwd,
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: `${join(cwd, "bin")}:${process.env.PATH}`,
        FAKE_RPM_LIST_FILE: fakeListFile,
        FAKE_RPM_REQUIRES_FILE: fakeRequiresFile,
        FAKE_RPM_MODE: fakeMode ?? "ok",
        FAKE_RPM_VERSION: fakeVersion ?? expectedVersion,
        FAKE_RPM_NAME: fakeName ?? expectedRpmName,
        ...(rpmNameOverride ? { FORMULA_RPM_NAME_OVERRIDE: rpmNameOverride } : {}),
      },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test(
  "validate-linux-rpm accepts an RPM whose file list contains the expected payload",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeRpmTool(binDir);
    writeFakeRpmExtractTools(binDir);

    // Fake RPM artifact (contents unused by the validator; it calls our fake rpm tool).
    writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

    const listFile = join(tmp, "rpm-list.txt");
    const requiresFile = writeDefaultRequiresFile(tmp);
    writeFileSync(
      listFile,
      [
        `/usr/bin/${expectedMainBinaryName}`,
        "/usr/share/applications/formula.desktop",
        `/usr/share/doc/${expectedRpmName}/LICENSE`,
        `/usr/share/doc/${expectedRpmName}/NOTICE`,
      ].join("\n"),
      { encoding: "utf8" },
    );

    // Run from tmp dir and pass a relative rpm path to ensure --rpm resolves against the invocation cwd.
    const proc = runValidator({
      cwd: tmp,
      rpmArg: "Formula.rpm",
      fakeListFile: listFile,
      fakeRequiresFile: requiresFile,
      fakeMode: "ok",
    });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test("validate-linux-rpm accepts --rpm pointing at a directory of RPMs", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula-1.rpm"), "not-a-real-rpm", { encoding: "utf8" });
  writeFileSync(join(tmp, "Formula-2.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinaryName}`,
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: ".", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-rpm fails when the expected desktop binary path is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(listFile, ["/usr/share/applications/formula.desktop"].join("\n"), { encoding: "utf8" });

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing expected desktop binary path/i);
});

test("validate-linux-rpm accepts when RPM %{NAME} is overridden for validation", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const overrideName = "formula-desktop-alt";
  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinaryName}`,
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${overrideName}/LICENSE`,
      `/usr/share/doc/${overrideName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeName: overrideName,
    rpmNameOverride: overrideName,
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-rpm fails when no .desktop file exists under /usr/share/applications/", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(listFile, ["/usr/bin/formula-desktop"].join("\n"), { encoding: "utf8" });

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing expected \.desktop file/i);
});

test("validate-linux-rpm fails when LICENSE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing compliance file/i);
  assert.match(proc.stderr, /LICENSE/i);
});

test("validate-linux-rpm fails when NOTICE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing compliance file/i);
  assert.match(proc.stderr, /NOTICE/i);
});

test("validate-linux-rpm fails when rpm --info query fails", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeMode: "fail-info",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rpm --info query failed/i);
});

test("validate-linux-rpm fails when rpm --queryformat fails", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeMode: "fail-queryformat",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rpm query failed for %\{VERSION\}/i);
});

test("validate-linux-rpm fails when RPM version does not match tauri.conf.json", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeVersion: "0.0.0",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /RPM version mismatch/i);
});

test("validate-linux-rpm fails when RPM name does not match tauri.conf.json", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeName: "some-other-name",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /RPM name mismatch/i);
});

test("validate-linux-rpm fails when extracted .desktop is missing MimeType=", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir, { withMimeType: false });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /No extracted \.desktop file contained a MimeType=/i);
});

test("validate-linux-rpm fails when extracted .desktop lacks xlsx integration", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  // Only advertise CSV (no xlsx substring + no canonical xlsx MIME).
  writeFakeRpmExtractTools(binDir, { mimeTypeLine: "MimeType=text/csv;" });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /advertised xlsx support/i);
});

test("validate-linux-rpm fails when extracted .desktop Exec= lacks a file placeholder", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir, { execLine: "Exec=formula-desktop" });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /placeholder/i);
});
