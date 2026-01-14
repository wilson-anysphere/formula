import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const tauriConfig = JSON.parse(
  readFileSync(join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"),
);
const expectedVersion = String(tauriConfig?.version ?? "").trim();

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
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
    withMimeTypeEntry = true,
    execLine = "formula-desktop %U",
    appImageVersion = expectedVersion,
  } = {},
) {
  const desktopMime = withXlsxMime
    ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;"
    : "text/plain;";

  const desktopBlock = withDesktopFile
    ? [
        "mkdir -p squashfs-root/usr/share/applications",
        "cat > squashfs-root/usr/share/applications/formula.desktop <<'DESKTOP'",
        "[Desktop Entry]",
        "Name=Formula",
        `Exec=${execLine}`,
        ...(appImageVersion ? [`X-AppImage-Version=${appImageVersion}`] : []),
        ...(withMimeTypeEntry ? [`MimeType=${desktopMime}`] : []),
        "DESKTOP",
      ].join("\n")
    : "mkdir -p squashfs-root/usr/share/applications";

  const script = `#!/usr/bin/env bash
set -euo pipefail

if [[ "\${1:-}" == "--appimage-extract" ]]; then
  mkdir -p squashfs-root/usr/bin
  mkdir -p squashfs-root/usr/share/doc/formula-desktop

  cat > squashfs-root/AppRun <<'APPRUN'
#!/usr/bin/env bash
echo "AppRun stub"
APPRUN
  chmod +x squashfs-root/AppRun

  cat > squashfs-root/usr/bin/formula-desktop <<'BIN'
#!/usr/bin/env bash
echo "formula-desktop stub"
BIN
  chmod +x squashfs-root/usr/bin/formula-desktop

  echo "LICENSE stub" > squashfs-root/usr/share/doc/formula-desktop/LICENSE
  echo "NOTICE stub" > squashfs-root/usr/share/doc/formula-desktop/NOTICE

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
  const proc = spawnSync(
    "bash",
    [join(repoRoot, "scripts", "validate-linux-appimage.sh"), "--appimage", appImagePath],
    {
      cwd: repoRoot,
      encoding: "utf8",
      env: {
        ...process.env,
        // Keep tests stable even if `mainBinaryName` changes or python isn't available.
        FORMULA_APPIMAGE_MAIN_BINARY: "formula-desktop",
      },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
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

test("validate-linux-appimage fails when .desktop lacks a MimeType entry", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-appimage-test-"));
  const appImagePath = join(tmp, "Formula.AppImage");
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withMimeTypeEntry: false, appImageVersion: expectedVersion });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /MimeType=/i);
});

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
  writeFakeAppImage(appImagePath, { withDesktopFile: true, withXlsxMime: true, execLine: "formula-desktop" });

  const proc = runValidator(appImagePath);
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Exec=.*placeholder|invalid Exec=/i);
});
