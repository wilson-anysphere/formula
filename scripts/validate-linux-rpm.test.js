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
const expectedRpmName = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";

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

if [[ "\${1:-}" != "-qp" ]]; then
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

  // `validate-linux-rpm.sh --no-container` extracts the RPM payload to validate
  // `.desktop` file MimeType entries, so provide lightweight stubs for rpm2cpio/cpio.
  const rpm2cpioPath = join(binDir, "rpm2cpio");
  writeFileSync(
    rpm2cpioPath,
    `#!/usr/bin/env bash
set -euo pipefail
# The real rpm2cpio converts an RPM to a CPIO archive. Our fake cpio ignores stdin,
# so we can just succeed without emitting anything.
exit 0
`,
    { encoding: "utf8" },
  );
  chmodSync(rpm2cpioPath, 0o755);

  const cpioPath = join(binDir, "cpio");
  writeFileSync(
    cpioPath,
    `#!/usr/bin/env bash
set -euo pipefail
# Create a minimal extracted payload tree expected by the validator.
mkdir -p usr/share/applications
cat > usr/share/applications/formula.desktop <<'DESKTOP'
[Desktop Entry]
Type=Application
Name=Formula
Exec=formula-desktop %U
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
DESKTOP
exit 0
`,
    { encoding: "utf8" },
  );
  chmodSync(cpioPath, 0o755);
}

function runValidator({ cwd, rpmArg, fakeListFile, fakeMode, fakeVersion, fakeName }) {
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
        FAKE_RPM_MODE: fakeMode ?? "ok",
        FAKE_RPM_VERSION: fakeVersion ?? expectedVersion,
        FAKE_RPM_NAME: fakeName ?? expectedRpmName,
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

    // Fake RPM artifact (contents unused by the validator; it calls our fake rpm tool).
    writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

    const listFile = join(tmp, "rpm-list.txt");
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

    // Run from tmp dir and pass a relative rpm path to ensure --rpm resolves against the invocation cwd.
    const proc = runValidator({
      cwd: tmp,
      rpmArg: "Formula.rpm",
      fakeListFile: listFile,
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

  writeFileSync(join(tmp, "Formula-1.rpm"), "not-a-real-rpm", { encoding: "utf8" });
  writeFileSync(join(tmp, "Formula-2.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
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

  const proc = runValidator({ cwd: tmp, rpmArg: ".", fakeListFile: listFile });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-rpm fails when /usr/bin/formula-desktop is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(listFile, ["/usr/share/applications/formula.desktop"].join("\n"), { encoding: "utf8" });

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing expected desktop binary path/i);
});

test("validate-linux-rpm fails when no .desktop file exists under /usr/share/applications/", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(listFile, ["/usr/bin/formula-desktop"].join("\n"), { encoding: "utf8" });

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile });
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
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile });
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
  writeFileSync(
    listFile,
    [
      "/usr/bin/formula-desktop",
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile });
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
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
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
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
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
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
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
  writeFileSync(
    listFile,
    ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop"].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeName: "some-other-name",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /RPM name mismatch/i);
});
