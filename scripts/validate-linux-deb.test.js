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
const expectedMainBinary = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";
const expectedIdentifier = String(tauriConf?.identifier ?? "").trim() || "app.formula.desktop";
const expectedMimeDefinitionContentsPath = `./usr/share/mime/packages/${expectedIdentifier}.xml`;
const expectedDebName = expectedMainBinary;
const expectedFileAssociationMimeTypes = Array.from(
  new Set(
    (tauriConf?.bundle?.fileAssociations ?? [])
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
const defaultDesktopMimeValue = `${expectedFileAssociationMimeTypes.join(";")};x-scheme-handler/formula;`;
const defaultDesktopMimeValueNoScheme = `${expectedFileAssociationMimeTypes.join(";")};`;

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

test("validate-linux-deb --help prints usage and mentions key env vars", { skip: !hasBash }, () => {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-linux-deb.sh"), "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /validate-linux-deb\.sh/i);
  assert.match(proc.stdout, /--no-container/);
  assert.match(proc.stdout, /DOCKER_PLATFORM/);
  assert.match(proc.stdout, /FORMULA_DEB_NAME_OVERRIDE/);
});

function writeFakeDpkgDebTool(binDir) {
  const script = `#!/usr/bin/env bash
 set -euo pipefail

 mode="\${FAKE_DPKG_DEB_MODE:-ok}"
 fake_version="\${FAKE_DPKG_DEB_VERSION:-0.0.0}"
 fake_package="\${FAKE_DPKG_DEB_PACKAGE:-formula-desktop}"
 fake_binary="\${FAKE_DPKG_DEB_BINARY:-formula-desktop}"

cmd="\${1:-}"
case "$cmd" in
  -f)
    deb="\${2:-}"
    field="\${3:-}"
    if [[ "$mode" == "fail-field" ]]; then
      echo "fake dpkg-deb: failing -f for $deb" >&2
      exit 1
    fi
    case "$field" in
      Version) echo "$fake_version" ;;
      Package) echo "$fake_package" ;;
      Depends)
        if [[ -z "\${FAKE_DPKG_DEB_DEPENDS_FILE:-}" ]]; then
          echo "fake dpkg-deb: missing FAKE_DPKG_DEB_DEPENDS_FILE" >&2
          exit 2
        fi
        cat "$FAKE_DPKG_DEB_DEPENDS_FILE"
        ;;
      *) echo "" ;;
    esac
    exit 0
    ;;
  -c|--contents)
    deb="\${2:-}"
    if [[ "$mode" == "fail-contents" ]]; then
      echo "fake dpkg-deb: failing -c for $deb" >&2
      exit 1
    fi
    if [[ -z "\${FAKE_DPKG_DEB_CONTENTS_FILE:-}" ]]; then
      echo "fake dpkg-deb: missing FAKE_DPKG_DEB_CONTENTS_FILE" >&2
      exit 2
    fi
    cat "$FAKE_DPKG_DEB_CONTENTS_FILE"
    exit 0
    ;;
  -x)
    deb="\${2:-}"
    dest="\${3:-}"
    if [[ "$mode" == "fail-extract" ]]; then
      echo "fake dpkg-deb: failing -x for $deb" >&2
      exit 1
    fi
    mkdir -p "$dest/usr/bin" "$dest/usr/share/applications" "$dest/usr/share/doc/$fake_package" "$dest/usr/share/mime/packages"
    cat > "$dest/usr/bin/$fake_package" <<'BIN'
#!/usr/bin/env bash
echo "formula stub"
BIN
    rm -f "$dest/usr/bin/$fake_package" || true
    cat > "$dest/usr/bin/$fake_binary" <<'BIN'
#!/usr/bin/env bash
echo "formula stub"
BIN
    chmod +x "$dest/usr/bin/$fake_binary"

    mime_value="\${FAKE_DESKTOP_MIME_VALUE:-${defaultDesktopMimeValue}}"
    exec_line="\${FAKE_DESKTOP_EXEC_LINE:-Exec=$fake_binary %U}"
    cat > "$dest/usr/share/applications/formula.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=Formula
\${exec_line}
MimeType=\${mime_value}
DESKTOP

    echo "LICENSE stub" > "$dest/usr/share/doc/$fake_package/LICENSE"
    echo "NOTICE stub" > "$dest/usr/share/doc/$fake_package/NOTICE"
    cat > "$dest/usr/share/mime/packages/${expectedIdentifier}.xml" <<'XML'
<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet">
    <glob pattern="*.parquet" />
  </mime-type>
</mime-info>
XML
    exit 0
    ;;
  *)
    echo "fake dpkg-deb: unsupported args: $*" >&2
    exit 2
    ;;
esac
`;

  const p = join(binDir, "dpkg-deb");
  writeFileSync(p, script, { encoding: "utf8" });
  chmodSync(p, 0o755);
}

function writeDefaultDependsFile(tmpDir) {
  const p = join(tmpDir, "deb-depends.txt");
  writeFileSync(
    p,
    [
      "shared-mime-info",
      "libwebkit2gtk-4.1-0",
      "libgtk-3-0",
      "libayatana-appindicator3-1",
      "librsvg2-2",
      "libssl3",
    ].join(", "),
    "utf8",
  );
  return p;
}

function writeDefaultContentsFile(
  tmpDir,
  {
    includeBinary = true,
    includeLicense = true,
    includeNotice = true,
    packageName = expectedDebName,
    binaryName = expectedMainBinary,
    includeParquetMimeDefinition = true,
  } = {},
) {
  const p = join(tmpDir, "deb-contents.txt");
  const lines = [];
  const add = (path) => lines.push(`-rw-r--r-- root/root 0 2024-01-01 00:00 ${path}`);
  if (includeBinary) add(`./usr/bin/${binaryName}`);
  add("./usr/share/applications/formula.desktop");
  if (includeLicense) add(`./usr/share/doc/${packageName}/LICENSE`);
  if (includeNotice) add(`./usr/share/doc/${packageName}/NOTICE`);
  if (includeParquetMimeDefinition) {
    add(expectedMimeDefinitionContentsPath);
  }
  writeFileSync(p, lines.join("\n"), "utf8");
  return p;
}

function runValidator({
  cwd,
  debArg,
  dependsFile,
  contentsFile,
  fakeMode,
  fakeVersion,
  fakePackage,
  fakeBinary,
  desktopMimeValue,
  desktopExecLine,
  debNameOverride,
} = {}) {
  const proc = spawnSync(
    "bash",
    [join(repoRoot, "scripts", "validate-linux-deb.sh"), "--no-container", "--deb", debArg],
    {
      cwd,
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: `${join(cwd, "bin")}:${process.env.PATH}`,
        FAKE_DPKG_DEB_MODE: fakeMode ?? "ok",
        FAKE_DPKG_DEB_VERSION: fakeVersion ?? expectedVersion,
        FAKE_DPKG_DEB_PACKAGE: fakePackage ?? expectedDebName,
        FAKE_DPKG_DEB_BINARY: fakeBinary ?? expectedMainBinary,
        FAKE_DPKG_DEB_DEPENDS_FILE: dependsFile,
        FAKE_DPKG_DEB_CONTENTS_FILE: contentsFile,
        ...(desktopMimeValue ? { FAKE_DESKTOP_MIME_VALUE: desktopMimeValue } : {}),
        ...(desktopExecLine ? { FAKE_DESKTOP_EXEC_LINE: desktopExecLine } : {}),
        ...(debNameOverride ? { FORMULA_DEB_NAME_OVERRIDE: debNameOverride } : {}),
      },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test(
  "validate-linux-deb accepts a DEB whose metadata + payload look correct",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeDpkgDebTool(binDir);

    writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
    const dependsFile = writeDefaultDependsFile(tmp);
    const contentsFile = writeDefaultContentsFile(tmp);

    const proc = runValidator({
      cwd: tmp,
      debArg: "Formula.deb",
      dependsFile,
      contentsFile,
    });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test("validate-linux-deb fails when the expected binary path is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp, { includeBinary: false });

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing expected desktop binary/i);
});

test("validate-linux-deb fails when LICENSE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp, { includeLicense: false });

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing compliance file/i);
  assert.match(proc.stderr, /LICENSE/i);
});

test("validate-linux-deb fails when NOTICE is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp, { includeNotice: false });

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing compliance file/i);
  assert.match(proc.stderr, /NOTICE/i);
});

test("validate-linux-deb accepts when DEB Version has a Debian revision suffix (-1)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    fakeVersion: `${expectedVersion}-1`,
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-deb fails when DEB Version uses a non-numeric suffix after the expected version", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    fakeVersion: `${expectedVersion}-beta.1`,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /DEB version mismatch/i);
});

test("validate-linux-deb accepts when Debian Package name is overridden for validation", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);
  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });

  const overrideName = "formula-desktop-alt";
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp, { packageName: overrideName, binaryName: expectedMainBinary });

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    fakePackage: overrideName,
    fakeBinary: expectedMainBinary,
    debNameOverride: overrideName,
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-deb accepts when extracted .desktop Exec= wraps the binary in quotes", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    desktopExecLine: `Exec=\"/usr/bin/${expectedMainBinary}\" %U`,
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-deb fails when shared-mime-info is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(dependsFile, ["libwebkit2gtk-4.1-0", "libgtk-3-0"].join(", "), "utf8");
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info/i);
});

test("validate-linux-deb fails when extracted .desktop lacks URL scheme handler (x-scheme-handler/formula)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    desktopMimeValue: defaultDesktopMimeValueNoScheme,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /x-scheme-handler\/formula/i);
});

test("validate-linux-deb fails when extracted .desktop lacks Parquet MIME type (application/vnd.apache.parquet)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const mimeTypesNoParquet = expectedFileAssociationMimeTypes.filter((mt) => mt !== "application/vnd.apache.parquet");
  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    desktopMimeValue: `${mimeTypesNoParquet.join(";")};x-scheme-handler/formula;`,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /application\/vnd\.apache\.parquet/i);
});

test("validate-linux-deb fails when extracted .desktop Exec= does not reference the expected binary", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({
    cwd: tmp,
    debArg: "Formula.deb",
    dependsFile,
    contentsFile,
    desktopExecLine: "Exec=something-else %U",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /target the expected executable/i);
});

test("validate-linux-deb fails when Parquet shared-mime-info definition is missing from the payload", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp, { includeParquetMimeDefinition: false });

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Parquet shared-mime-info/i);
});
