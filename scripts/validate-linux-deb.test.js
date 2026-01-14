import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
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

function collectDeepLinkSchemes(config) {
  const deepLink = config?.plugins?.["deep-link"];
  const desktop = deepLink?.desktop;
  const schemes = new Set();
  const addFromProtocol = (protocol) => {
    if (!protocol || typeof protocol !== "object") return;
    const raw = protocol.schemes;
    const values = typeof raw === "string" ? [raw] : Array.isArray(raw) ? raw : [];
    for (const v of values) {
      if (typeof v !== "string") continue;
      const normalized = v.trim().replace(/[:/]+$/, "").toLowerCase();
      if (normalized) schemes.add(normalized);
    }
  };
  if (Array.isArray(desktop)) {
    for (const protocol of desktop) addFromProtocol(protocol);
  } else {
    addFromProtocol(desktop);
  }
  if (schemes.size === 0) schemes.add("formula");
  return Array.from(schemes).sort();
}

const expectedDeepLinkSchemes = collectDeepLinkSchemes(tauriConf);
const expectedSchemeMimes = expectedDeepLinkSchemes.map((scheme) => `x-scheme-handler/${scheme}`);
const defaultDesktopMimeValue = `${expectedFileAssociationMimeTypes.join(";")};${expectedSchemeMimes.join(";")};`;
const defaultDesktopMimeValueNoScheme = `${expectedFileAssociationMimeTypes.join(";")};`;

function buildSharedMimeInfoXml({ omitGlobsForExts = new Set() } = {}) {
  const groups = new Map();
  const associations = Array.isArray(tauriConf?.bundle?.fileAssociations) ? tauriConf.bundle.fileAssociations : [];
  for (const assoc of associations) {
    const mimeType = typeof assoc?.mimeType === "string" ? assoc.mimeType.trim() : "";
    if (!mimeType) continue;
    const rawExts = assoc?.ext;
    const exts = Array.isArray(rawExts) ? rawExts : typeof rawExts === "string" ? [rawExts] : [];
    for (const raw of exts) {
      if (typeof raw !== "string") continue;
      const ext = raw.trim().replace(/^\./, "").toLowerCase();
      if (!ext) continue;
      if (!groups.has(mimeType)) groups.set(mimeType, new Set());
      groups.get(mimeType).add(ext);
    }
  }

  const lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
  ];
  for (const mimeType of Array.from(groups.keys()).sort()) {
    lines.push(`  <mime-type type="${mimeType}">`);
    const exts = Array.from(groups.get(mimeType)).sort();
    for (const ext of exts) {
      if (omitGlobsForExts.has(ext)) continue;
      lines.push(`    <glob pattern="*.${ext}" />`);
    }
    lines.push("  </mime-type>");
  }
  lines.push("</mime-info>");
  return lines.join("\n");
}

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
  assert.match(proc.stdout, /FORMULA_TAURI_CONF_PATH/);
  assert.match(proc.stdout, /FORMULA_DEB_NAME_OVERRIDE/);
});

test("validate-linux-deb bounds extracted .desktop discovery to avoid deep scans (perf guardrail)", () => {
  const script = readFileSync(join(repoRoot, "scripts", "validate-linux-deb.sh"), "utf8");
  const idx = script.indexOf('find "$applications_dir"');
  assert.ok(idx >= 0, "Expected validate-linux-deb.sh to use find \"$applications_dir\" when validating extracted desktop entries.");
  const snippet = script.slice(idx, idx + 200);
  assert.ok(
    snippet.includes("-maxdepth"),
    `Expected validate-linux-deb.sh to bound the .desktop scan depth with -maxdepth.\nSaw snippet:\n${snippet}`,
  );
});

function writeFakeDpkgDebTool(binDir) {
  const defaultMimeXml = buildSharedMimeInfoXml();
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
    with_mimetype="\${FAKE_DESKTOP_WITH_MIMETYPE:-1}"
    {
      echo "[Desktop Entry]"
      echo "Type=Application"
      echo "Name=Formula"
      echo "\${exec_line}"
      if [[ "\${with_mimetype}" == "1" ]]; then
        echo "MimeType=\${mime_value}"
      fi
    } > "$dest/usr/share/applications/formula.desktop"

    echo "LICENSE stub" > "$dest/usr/share/doc/$fake_package/LICENSE"
    echo "NOTICE stub" > "$dest/usr/share/doc/$fake_package/NOTICE"
    if [[ -n "\${FAKE_MIME_XML_CONTENT:-}" ]]; then
      printf '%s\n' "$FAKE_MIME_XML_CONTENT" > "$dest/usr/share/mime/packages/${expectedIdentifier}.xml"
    else
      cat > "$dest/usr/share/mime/packages/${expectedIdentifier}.xml" <<'XML'
${defaultMimeXml}
XML
    fi
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

function writeFakePython3Tool(binDir) {
  // The DEB validator optionally runs scripts/ci/verify_linux_desktop_integration.py (python)
  // after its bash-based checks. For certain tests we want to validate the bash fallback
  // logic without the python verifier masking results, so we stub python3 accordingly.
  const script = `#!/usr/bin/env bash
set -euo pipefail

if [[ "\${1:-}" == "-" ]]; then
  cat >/dev/null || true
  key="\${3:-}"
  case "$key" in
    version) printf '%s\\n' "${expectedVersion}" ;;
    mainBinaryName) printf '%s\\n' "${expectedMainBinary}" ;;
    identifier) printf '%s\\n' "${expectedIdentifier}" ;;
  esac
  exit 0
fi

if [[ "\${1:-}" == *"verify_linux_desktop_integration.py" ]]; then
  exit 0
fi

exit 0
`;

  const pythonPath = join(binDir, "python3");
  writeFileSync(pythonPath, script, { encoding: "utf8" });
  chmodSync(pythonPath, 0o755);
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
  desktopWithMimeType,
  mimeXmlContent,
  debNameOverride,
  tauriConfPath,
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
        ...(desktopWithMimeType === false ? { FAKE_DESKTOP_WITH_MIMETYPE: "0" } : {}),
        ...(mimeXmlContent ? { FAKE_MIME_XML_CONTENT: mimeXmlContent } : {}),
        ...(desktopMimeValue ? { FAKE_DESKTOP_MIME_VALUE: desktopMimeValue } : {}),
        ...(desktopExecLine ? { FAKE_DESKTOP_EXEC_LINE: desktopExecLine } : {}),
        ...(debNameOverride ? { FORMULA_DEB_NAME_OVERRIDE: debNameOverride } : {}),
        ...(tauriConfPath ? { FORMULA_TAURI_CONF_PATH: tauriConfPath } : {}),
      },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test("validate-linux-deb honors FORMULA_TAURI_CONF_PATH (relative to repo root)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const overrideVersion = "0.0.0";
  const confParent = join(repoRoot, "target");
  mkdirSync(confParent, { recursive: true });
  const confDir = mkdtempSync(join(confParent, "tauri-conf-override-"));
  const confPath = join(confDir, "tauri.conf.json");
  writeFileSync(confPath, JSON.stringify({ ...tauriConf, version: overrideVersion }), { encoding: "utf8" });

  try {
    const proc = runValidator({
      cwd: tmp,
      debArg: "Formula.deb",
      dependsFile,
      contentsFile,
      fakeVersion: overrideVersion,
      tauriConfPath: relative(repoRoot, confPath),
    });
    assert.equal(proc.status, 0, proc.stderr);
  } finally {
    rmSync(confDir, { recursive: true, force: true });
  }
});

test("validate-linux-deb rejects tauri identifiers containing path separators", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = writeDefaultDependsFile(tmp);
  const contentsFile = writeDefaultContentsFile(tmp);

  const confParent = join(repoRoot, "target");
  mkdirSync(confParent, { recursive: true });
  const confDir = mkdtempSync(join(confParent, "tauri-conf-override-"));
  const confPath = join(confDir, "tauri.conf.json");
  writeFileSync(confPath, JSON.stringify({ ...tauriConf, identifier: "com/example.formula.desktop" }), {
    encoding: "utf8",
  });

  try {
    const proc = runValidator({
      cwd: tmp,
      debArg: "Formula.deb",
      dependsFile,
      contentsFile,
      tauriConfPath: relative(repoRoot, confPath),
    });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identifier.*valid filename/i);
    assert.match(proc.stderr, /path separators/i);
  } finally {
    rmSync(confDir, { recursive: true, force: true });
  }
});

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

test("validate-linux-deb fails when WebKitGTK 4.1 dependency is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(
    dependsFile,
    [
      "shared-mime-info",
      // Deliberately declare WebKitGTK 4.0 (should reject; we require 4.1).
      "libwebkit2gtk-4.0-37",
      "libgtk-3-0",
      "libayatana-appindicator3-1",
      "librsvg2-2",
      "libssl3",
    ].join(", "),
    "utf8",
  );
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /WebKitGTK 4\.1/i);
});

test("validate-linux-deb fails when OpenSSL (libssl) dependency is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(
    dependsFile,
    [
      "shared-mime-info",
      "libwebkit2gtk-4.1-0",
      "libgtk-3-0",
      "libayatana-appindicator3-1",
      "librsvg2-2",
      // Deliberately omit libssl dependency.
    ].join(", "),
    "utf8",
  );
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /OpenSSL/i);
});

test("validate-linux-deb fails when GTK3 dependency is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(
    dependsFile,
    [
      "shared-mime-info",
      "libwebkit2gtk-4.1-0",
      // Deliberately declare GTK4 instead of GTK3.
      "libgtk-4-1",
      "libayatana-appindicator3-1",
      "librsvg2-2",
      "libssl3",
    ].join(", "),
    "utf8",
  );
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /GTK3/i);
});

test("validate-linux-deb fails when AppIndicator dependency is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(
    dependsFile,
    [
      "shared-mime-info",
      "libwebkit2gtk-4.1-0",
      "libgtk-3-0",
      // Deliberately omit libappindicator/libayatana-appindicator.
      "librsvg2-2",
      "libssl3",
    ].join(", "),
    "utf8",
  );
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /AppIndicator/i);
});

test("validate-linux-deb fails when librsvg dependency is missing from Depends", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeDpkgDebTool(binDir);

  writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
  const dependsFile = join(tmp, "deb-depends.txt");
  writeFileSync(
    dependsFile,
    [
      "shared-mime-info",
      "libwebkit2gtk-4.1-0",
      "libgtk-3-0",
      "libayatana-appindicator3-1",
      // Deliberately omit librsvg2-2.
      "libssl3",
    ].join(", "),
    "utf8",
  );
  const contentsFile = writeDefaultContentsFile(tmp);

  const proc = runValidator({ cwd: tmp, debArg: "Formula.deb", dependsFile, contentsFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /librsvg/i);
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

test(
  "validate-linux-deb requires URL scheme handlers to match exact MimeType= tokens (no prefix matches)",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-deb-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeDpkgDebTool(binDir);
    writeFakePython3Tool(binDir);

    writeFileSync(join(tmp, "Formula.deb"), "not-a-real-deb", { encoding: "utf8" });
    const dependsFile = writeDefaultDependsFile(tmp);
    const contentsFile = writeDefaultContentsFile(tmp);

    const prefixSchemeMimes = expectedSchemeMimes.map((schemeMime) => `${schemeMime}-extra`);
    const proc = runValidator({
      cwd: tmp,
      debArg: "Formula.deb",
      dependsFile,
      contentsFile,
      desktopMimeValue: `${expectedFileAssociationMimeTypes.join(";")};${prefixSchemeMimes.join(";")};`,
    });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /expected URL scheme handler/i);
  },
);

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

test("validate-linux-deb fails when extracted .desktop Exec= lacks a file/URL placeholder", { skip: !hasBash }, () => {
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
    desktopExecLine: `Exec=${expectedMainBinary}`,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /placeholder/i);
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

test("validate-linux-deb fails when Parquet shared-mime-info definition is missing expected content", { skip: !hasBash }, () => {
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
    mimeXmlContent: `<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet" />
</mime-info>`,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  // When python3 is present, the strict verifier reports no packaged shared-mime-info definition
  // matching the expected MIME + glob; otherwise the bash fallback emits a "missing expected content"
  // error. Match either so the test remains hermetic across environments.
  assert.match(
    proc.stderr,
    /(missing expected content|missing required content|missing required glob mappings|no packaged shared-mime-info definition)/i,
  );
});

test("validate-linux-deb fails when extracted .desktop is missing MimeType=", { skip: !hasBash }, () => {
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
    desktopWithMimeType: false,
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /MimeType=/i);
});

test("validate-linux-deb fails when extracted .desktop lacks xlsx integration", { skip: !hasBash }, () => {
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
    // Only advertise csv + parquet (no xlsx MIME types).
    desktopMimeValue: "text/csv;application/vnd.apache.parquet;x-scheme-handler/formula;",
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  // python verifier complains about missing spreadsheetml.sheet; bash fallback complains about xlsx support.
  assert.match(proc.stderr, /(xlsx|spreadsheetml\.sheet)/i);
});
