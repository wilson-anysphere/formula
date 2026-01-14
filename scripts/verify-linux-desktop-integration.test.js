import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "verify_linux_desktop_integration.py");

const hasPython3 = (() => {
  const probe = spawnSync("python3", ["--version"], { stdio: "ignore" });
  return !probe.error && probe.status === 0;
})();

function writeConfig(dir, { mainBinaryName = "formula-desktop" } = {}) {
  const configPath = path.join(dir, "tauri.conf.json");
  const conf = {
    mainBinaryName,
    bundle: {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      ],
    },
  };
  writeFileSync(configPath, JSON.stringify(conf), "utf8");
  return configPath;
}

function writePackageRoot(dir, { execLine, mimeTypeLine, docPackageName = "formula-desktop" } = {}) {
  const pkgRoot = path.join(dir, "pkg");
  const applicationsDir = path.join(pkgRoot, "usr", "share", "applications");
  const docDir = path.join(pkgRoot, "usr", "share", "doc", docPackageName);
  mkdirSync(applicationsDir, { recursive: true });
  mkdirSync(docDir, { recursive: true });
  writeFileSync(path.join(docDir, "LICENSE"), "stub", "utf8");
  writeFileSync(path.join(docDir, "NOTICE"), "stub", "utf8");

  const desktopPath = path.join(applicationsDir, "formula.desktop");
  const exec = execLine ?? "formula-desktop %U";
  const mime =
    mimeTypeLine ??
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;";
  const desktop = [
    "[Desktop Entry]",
    "Name=Formula",
    `Exec=${exec}`,
    `MimeType=${mime}`,
  ].join("\n");
  writeFileSync(desktopPath, desktop, "utf8");
  return pkgRoot;
}

function runValidator({ packageRoot, configPath, extraArgs = [] }) {
  const proc = spawnSync(
    "python3",
    [scriptPath, "--package-root", packageRoot, "--tauri-config", configPath, ...extraArgs],
    { cwd: repoRoot, encoding: "utf8" },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test("verify_linux_desktop_integration passes for a desktop entry targeting the expected binary", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp);

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify_linux_desktop_integration fails when no .desktop entries target the expected binary", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp, { execLine: "something-else %U" });

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /target the expected executable/i);
});

test(
  "verify_linux_desktop_integration supports overriding doc package name without changing Exec binary target",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfig(tmp, { mainBinaryName: "formula-desktop" });
    const pkgRoot = writePackageRoot(tmp, { docPackageName: "formula-desktop-alt", execLine: "formula-desktop %U" });

    const proc = runValidator({
      packageRoot: pkgRoot,
      configPath,
      extraArgs: ["--doc-package-name", "formula-desktop-alt", "--expected-main-binary", "formula-desktop"],
    });
    assert.equal(proc.status, 0, proc.stderr);
  },
);
