import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../apps/desktop/test/sourceTextUtils.js";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const validatorScriptPath = join(repoRoot, "scripts", "validate-linux-rpm.sh");
const validatorScriptContents = stripHashComments(readFileSync(validatorScriptPath, "utf8"));
const tauriConf = JSON.parse(readFileSync(join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"));
const expectedVersion = String(tauriConf?.version ?? "").trim();
const expectedMainBinary = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";
const expectedIdentifier = String(tauriConf?.identifier ?? "").trim() || "app.formula.desktop";
const expectedMimeDefinitionPath = `/usr/share/mime/packages/${expectedIdentifier}.xml`;
const expectedRpmName = expectedMainBinary;
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
const defaultDesktopMimeValue = `${expectedFileAssociationMimeTypes.join(";")};`;

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

test("validate-linux-rpm.sh bounds fallback desktop file scans (perf guardrail)", () => {
  const lines = validatorScriptContents.split(/\r?\n/);
  let found = false;
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const trimmed = raw.trim();
    if (!trimmed) continue;
    if (!raw.includes('find "$tmpdir"')) continue;
    found = true;

    const snippet = lines.slice(i, i + 12).join("\n");
    assert.match(
      snippet,
      /-maxdepth\s+\d+/,
      `Expected find \"$tmpdir\" fallback scan to be bounded with -maxdepth.\nSaw snippet:\n${snippet}`,
    );
  }
  assert.ok(found, 'Expected validate-linux-rpm.sh to include a find "$tmpdir" fallback scan for desktop entries.');
});

test("validate-linux-rpm.sh bounds extracted .desktop discovery to avoid deep scans (perf guardrail)", () => {
  for (const needle of ['find "$applications_dir"', 'find "$alt_dir"']) {
    const idx = validatorScriptContents.indexOf(needle);
    assert.ok(idx >= 0, `Expected validate-linux-rpm.sh to include ${needle} for desktop entry validation.`);
    const snippet = validatorScriptContents.slice(idx, idx + 200);
    assert.ok(
      snippet.includes("-maxdepth"),
      `Expected ${needle} scan to include -maxdepth.\nSaw snippet:\n${snippet}`,
    );
  }
});

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

test("validate-linux-rpm --help prints usage and mentions key env vars", { skip: !hasBash }, () => {
  const proc = spawnSync("bash", [join(repoRoot, "scripts", "validate-linux-rpm.sh"), "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /validate-linux-rpm\.sh/i);
  assert.match(proc.stdout, /--no-container/);
  assert.match(proc.stdout, /DOCKER_PLATFORM/);
  assert.match(proc.stdout, /FORMULA_TAURI_CONF_PATH/);
  assert.match(proc.stdout, /FORMULA_RPM_NAME_OVERRIDE/);
});

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

function writeFakePython3Tool(binDir) {
  // Many CI environments have python3 available, which causes the RPM validator to run
  // scripts/ci/verify_linux_desktop_integration.py after its bash-based checks.
  //
  // For certain guardrail tests, we want to exercise the bash fallback logic *without*
  // the python verifier masking failures/successes. This stub returns expected values
  // for the tauri.conf.json lookups and treats the python verifier invocation as a no-op.
  const script = `#!/usr/bin/env bash
set -euo pipefail

if [[ "\${1:-}" == "-" ]]; then
  # Drain stdin (the validator passes python source via heredoc).
  cat >/dev/null || true
  key="\${3:-}"
  case "$key" in
    version) printf '%s\\n' "${expectedVersion}" ;;
    mainBinaryName) printf '%s\\n' "${expectedMainBinary}" ;;
    identifier) printf '%s\\n' "${expectedIdentifier}" ;;
  esac
  exit 0
fi

# Skip running the real verifier so tests can validate the bash scheme-token logic in isolation.
if [[ "\${1:-}" == *"verify_linux_desktop_integration.py" ]]; then
  exit 0
fi

exit 0
`;
  const pythonPath = join(binDir, "python3");
  writeFileSync(pythonPath, script, { encoding: "utf8" });
  chmodSync(pythonPath, 0o755);
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
    mimeTypeLine = `MimeType=${defaultDesktopMimeValue}`,
    withSchemeMime = true,
    withParquetMime = true,
    execLine = `Exec=${expectedMainBinary} %U`,
    docPackageName = expectedRpmName,
    mimeXmlContent = buildSharedMimeInfoXml(),
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
  if (withMimeType) {
    // Parse the provided MimeType= line into a stable list so we can reliably
    // include/exclude specific entries regardless of what tauri.conf.json
    // currently advertises.
    const raw = effectiveMimeTypeLine.replace(/^MimeType=/i, "");
    const tokens = raw
      .split(";")
      .map((t) => t.trim())
      .filter(Boolean);

    const hasToken = (value) => tokens.some((t) => t.toLowerCase() === value.toLowerCase());
    const removeToken = (value) => {
      const lower = value.toLowerCase();
      for (let i = tokens.length - 1; i >= 0; i -= 1) {
        if (tokens[i].toLowerCase() === lower) tokens.splice(i, 1);
      }
    };

    // Tests pass flags like `withParquetMime: false` to ensure the extracted
    // desktop file lacks that MIME type, even if it is included in the default
    // fileAssociation set.
    if (!withParquetMime) removeToken("application/vnd.apache.parquet");
    if (!withSchemeMime) {
      for (const schemeMime of expectedSchemeMimes) removeToken(schemeMime);
    }

    if (withParquetMime && !hasToken("application/vnd.apache.parquet")) tokens.push("application/vnd.apache.parquet");
    if (withSchemeMime) {
      for (const schemeMime of expectedSchemeMimes) {
        if (!hasToken(schemeMime)) tokens.push(schemeMime);
      }
    }

    effectiveMimeTypeLine = `MimeType=${tokens.join(";")};`;
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

mkdir -p usr/share/applications usr/share/mime/packages
cat > usr/share/applications/formula.desktop <<'DESKTOP'
${desktopLines.join("\n")}
DESKTOP

mkdir -p usr/share/doc/${docPackageName}
echo "LICENSE stub" > usr/share/doc/${docPackageName}/LICENSE
echo "NOTICE stub" > usr/share/doc/${docPackageName}/NOTICE

cat > usr/share/mime/packages/${expectedIdentifier}.xml <<'XML'
${mimeXmlContent}
XML
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
  tauriConfPath,
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
        ...(tauriConfPath ? { FORMULA_TAURI_CONF_PATH: tauriConfPath } : {}),
      },
    },
  );
  if (proc.error) throw proc.error;
  return proc;
}

test("validate-linux-rpm honors FORMULA_TAURI_CONF_PATH (relative to repo root)", { skip: !hasBash }, () => {
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
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const overrideVersion = "0.0.0";
  const confParent = join(repoRoot, "target");
  mkdirSync(confParent, { recursive: true });
  const confDir = mkdtempSync(join(confParent, "tauri-conf-override-"));
  const confPath = join(confDir, "tauri.conf.json");
  writeFileSync(confPath, JSON.stringify({ ...tauriConf, version: overrideVersion }), { encoding: "utf8" });

  try {
    const proc = runValidator({
      cwd: tmp,
      rpmArg: "Formula.rpm",
      fakeListFile: listFile,
      fakeRequiresFile: requiresFile,
      fakeVersion: overrideVersion,
      tauriConfPath: relative(repoRoot, confPath),
    });
    assert.equal(proc.status, 0, proc.stderr);
  } finally {
    rmSync(confDir, { recursive: true, force: true });
  }
});

test("validate-linux-rpm rejects tauri identifiers containing path separators", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);

  // Fake RPM artifact (the validator should fail before it inspects it).
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(listFile, "", { encoding: "utf8" });
  const requiresFile = join(tmp, "rpm-requires.txt");
  writeFileSync(requiresFile, "", { encoding: "utf8" });

  const confParent = join(repoRoot, "target");
  mkdirSync(confParent, { recursive: true });
  const confDir = mkdtempSync(join(confParent, "tauri-conf-override-"));
  const confPath = join(confDir, "tauri.conf.json");
  writeFileSync(confPath, JSON.stringify({ ...tauriConf, identifier: "com/example.formula.desktop" }), { encoding: "utf8" });

  try {
    const proc = runValidator({
      cwd: tmp,
      rpmArg: "Formula.rpm",
      fakeListFile: listFile,
      fakeRequiresFile: requiresFile,
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
        `/usr/bin/${expectedMainBinary}`,
        "/usr/share/applications/formula.desktop",
        expectedMimeDefinitionPath,
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

test("validate-linux-rpm accepts when extracted .desktop Exec= wraps the binary in quotes", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir, { execLine: `Exec="/usr/bin/${expectedMainBinary}" %U` });

  // Fake RPM artifact (contents unused by the validator; it calls our fake rpm tool).
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({
    cwd: tmp,
    rpmArg: "Formula.rpm",
    fakeListFile: listFile,
    fakeRequiresFile: requiresFile,
    fakeMode: "ok",
  });
  assert.equal(proc.status, 0, proc.stderr);
});

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
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: ".", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.equal(proc.status, 0, proc.stderr);
});

test("validate-linux-rpm fails when extracted .desktop Exec= does not reference the expected binary", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir, { execLine: "Exec=something-else %U" });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /target the expected executable/i);
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
  const overrideName = "formula-desktop-alt";
  writeFakeRpmExtractTools(binDir, { docPackageName: overrideName });
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });
  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
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
  writeFileSync(listFile, [`/usr/bin/${expectedMainBinary}`].join("\n"), { encoding: "utf8" });

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
      `/usr/bin/${expectedMainBinary}`,
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
      `/usr/bin/${expectedMainBinary}`,
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

test("validate-linux-rpm fails when Parquet shared-mime-info definition is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  // The validator still attempts an extraction-based desktop integration check when --no-container is
  // used, even if static validation already flagged missing payload entries. Provide fake extraction
  // tools so this test does not depend on system rpm2cpio/cpio availability.
  writeFakeRpmExtractTools(binDir);
  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /Parquet shared-mime-info/i);
});

test("validate-linux-rpm fails when Parquet shared-mime-info definition is missing expected content", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir, {
    // Deliberately omit the '*.parquet' glob so the validator fails the extraction content check.
    mimeXmlContent: `<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet" />
</mime-info>`,
  });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info definition file is missing expected content/i);
});

test("validate-linux-rpm fails when shared-mime-info is missing from RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-missing-shared-mime-info.txt");
  writeFileSync(
    requiresFile,
    [
      // Deliberately omit `shared-mime-info`.
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing required dependency/i);
  assert.match(proc.stderr, /shared-mime-info/i);
});

test("validate-linux-rpm fails when GTK3 rich OR dependency is missing from RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-missing-gtk3.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      // Deliberately omit GTK3 dependency.
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rich dependency OR expression/i);
  assert.match(proc.stderr, /GTK3/i);
});

test("validate-linux-rpm fails when AppIndicator rich OR dependency is missing from RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-missing-appindicator.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      // Deliberately omit AppIndicator/Ayatana dependency.
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rich dependency OR expression/i);
  assert.match(proc.stderr, /AppIndicator/i);
});

test("validate-linux-rpm fails when librsvg rich OR dependency is missing from RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-missing-librsvg.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      // Deliberately omit librsvg dependency.
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rich dependency OR expression/i);
  assert.match(proc.stderr, /librsvg/i);
});

test("validate-linux-rpm fails when WebKitGTK dependency is not expressed as a rich OR in RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-bad-webkit.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      // Deliberately list both distro alternatives as separate Requires lines, instead of
      // a single RPM rich dependency expression ("(A or B)"). This would make the RPM
      // uninstallable on at least one distro family.
      "webkit2gtk4.1",
      "libwebkit2gtk-4_1-0",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rich dependency OR expression/i);
  assert.match(proc.stderr, /WebKitGTK 4\.1/i);
});

test("validate-linux-rpm fails when RPM Requires reference WebKitGTK 4.0 instead of 4.1", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-webkit-4.0.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      // Deliberately use 4.0 package names (should reject; we require 4.1).
      "(webkit2gtk4.0 or libwebkit2gtk-4_0-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /WebKitGTK 4\.1/i);
});

test("validate-linux-rpm fails when OpenSSL rich OR dependency is missing from RPM Requires", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  writeFakeRpmExtractTools(binDir);

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const requiresFile = join(tmp, "rpm-requires-missing-openssl.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      // Deliberately omit OpenSSL rich dependency.
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /rich dependency OR expression/i);
  assert.match(proc.stderr, /OpenSSL/i);
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
    [`/usr/bin/${expectedMainBinary}`, "/usr/share/applications/formula.desktop"].join("\n"),
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
    [`/usr/bin/${expectedMainBinary}`, "/usr/share/applications/formula.desktop"].join("\n"),
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
    [`/usr/bin/${expectedMainBinary}`, "/usr/share/applications/formula.desktop"].join("\n"),
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
    [`/usr/bin/${expectedMainBinary}`, "/usr/share/applications/formula.desktop"].join("\n"),
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
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
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
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
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
  writeFakeRpmExtractTools(binDir, { execLine: `Exec=${expectedMainBinary}` });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /placeholder/i);
});

test("validate-linux-rpm fails when extracted .desktop lacks Parquet MIME type (application/vnd.apache.parquet)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  const mimeTypesNoParquet = expectedFileAssociationMimeTypes.filter((mt) => mt !== "application/vnd.apache.parquet");
  writeFakeRpmExtractTools(binDir, {
    withParquetMime: false,
    mimeTypeLine: `MimeType=${mimeTypesNoParquet.join(";")};`,
  });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /parquet/i);
});

test("validate-linux-rpm fails when extracted .desktop lacks URL scheme handler (x-scheme-handler/formula)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
  const binDir = join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeRpmTool(binDir);
  // Only advertise xlsx MIME type (no x-scheme-handler/formula).
  writeFakeRpmExtractTools(binDir, {
    mimeTypeLine: "MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;",
    withSchemeMime: false,
  });

  writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  const listFile = join(tmp, "rpm-list.txt");
  const requiresFile = writeDefaultRequiresFile(tmp);
  writeFileSync(
    listFile,
    [
      `/usr/bin/${expectedMainBinary}`,
      "/usr/share/applications/formula.desktop",
      expectedMimeDefinitionPath,
      `/usr/share/doc/${expectedRpmName}/LICENSE`,
      `/usr/share/doc/${expectedRpmName}/NOTICE`,
    ].join("\n"),
    { encoding: "utf8" },
  );

  const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /x-scheme-handler\/formula/i);
});

test(
  "validate-linux-rpm requires URL scheme handlers to match exact MimeType= tokens (no prefix matches)",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(join(tmpdir(), "formula-rpm-test-"));
    const binDir = join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeRpmTool(binDir);
    writeFakePython3Tool(binDir);

    const prefixSchemeMimes = expectedSchemeMimes.map((schemeMime) => `${schemeMime}-extra`);
    const mimeTypeLine = `MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;${prefixSchemeMimes.join(";")};`;
    writeFakeRpmExtractTools(binDir, { mimeTypeLine, withSchemeMime: false });

    writeFileSync(join(tmp, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

    const listFile = join(tmp, "rpm-list.txt");
    const requiresFile = writeDefaultRequiresFile(tmp);
    writeFileSync(
      listFile,
      [
        `/usr/bin/${expectedMainBinary}`,
        "/usr/share/applications/formula.desktop",
        expectedMimeDefinitionPath,
        `/usr/share/doc/${expectedRpmName}/LICENSE`,
        `/usr/share/doc/${expectedRpmName}/NOTICE`,
      ].join("\n"),
      { encoding: "utf8" },
    );

    const proc = runValidator({ cwd: tmp, rpmArg: "Formula.rpm", fakeListFile: listFile, fakeRequiresFile: requiresFile });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /expected URL scheme handler/i);
  },
);
