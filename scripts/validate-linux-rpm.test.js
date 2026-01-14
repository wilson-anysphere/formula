import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

function writeFakeRpmTool(binDir) {
  const rpmScript = `#!/usr/bin/env bash
set -euo pipefail

mode="\${FAKE_RPM_MODE:-ok}"

if [[ "\${1:-}" != "-qp" ]]; then
  echo "fake rpm: unexpected args: $*" >&2
  exit 2
fi

query="\${2:-}"
rpm_path="\${3:-}"

case "$query" in
  --info)
    if [[ "$mode" == "fail-info" ]]; then
      echo "fake rpm: failing --info for $rpm_path" >&2
      exit 1
    fi
    echo "Name        : formula-desktop"
    echo "Version     : 0.0.0"
    exit 0
    ;;
  --list)
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

function runValidator({ cwd, rpmArg, fakeListFile, fakeMode }) {
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
      ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop", "/usr/share/doc/formula/README"].join(
        "\n",
      ),
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

test(
  "validate-linux-rpm accepts --rpm pointing at a directory of RPMs",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeRpmTool(binDir);

    writeFileSync(join(tmp, "Formula-1.rpm"), "not-a-real-rpm", { encoding: "utf8" });
    writeFileSync(join(tmp, "Formula-2.rpm"), "not-a-real-rpm", { encoding: "utf8" });

    const listFile = join(tmp, "rpm-list.txt");
    writeFileSync(
      listFile,
      ["/usr/bin/formula-desktop", "/usr/share/applications/formula.desktop", "/usr/share/doc/formula/README"].join(
        "\n",
      ),
      { encoding: "utf8" },
    );

    // Run from tmp dir and pass a relative directory to ensure --rpm resolves against the invocation cwd.
    const proc = runValidator({ cwd: tmp, rpmArg: ".", fakeListFile: listFile, fakeMode: "ok" });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "validate-linux-rpm fails when /usr/bin/formula-desktop is missing",
  { skip: !hasBash },
  () => {
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
  },
);

test(
  "validate-linux-rpm fails when no .desktop file exists under /usr/share/applications/",
  { skip: !hasBash },
  () => {
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
  },
);

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
