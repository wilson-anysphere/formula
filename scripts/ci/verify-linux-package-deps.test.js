import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const tauriConf = JSON.parse(readFileSync(path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"), "utf8"));

const expectedVersion = String(tauriConf?.version ?? "").trim();
const expectedMainBinary = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";
const expectedIdentifier = String(tauriConf?.identifier ?? "").trim() || "app.formula.desktop";

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

test("verify-linux-package-deps bounds fallback package discovery scans (perf guardrail)", () => {
  const script = stripHashComments(readFileSync(path.join(repoRoot, "scripts", "ci", "verify-linux-package-deps.sh"), "utf8"));
  let found = false;
  for (const needle of ['find "${bundle_dirs[@]}"', 'find \"${bundle_dirs[@]}\"']) {
    // There should be at least one find invocation using the bundle_dirs array.
    if (!script.includes(needle)) continue;
    found = true;
    const idx = script.indexOf(needle);
    const snippet = script.slice(idx, idx + 200);
    assert.ok(
      snippet.includes("-maxdepth"),
      `Expected verify-linux-package-deps.sh to bound fallback find scans over bundle dirs with -maxdepth.\nSaw snippet:\n${snippet}`,
    );
  }
  assert.ok(found, "Expected verify-linux-package-deps.sh to use find over the bundle_dirs array when discovering packages.");
  // Ensure we don't regress to the previous unbounded form.
  assert.doesNotMatch(script, /find \"\\$\\{bundle_dirs\\[@\\]\\}\" -type f -name \"\\*\\.deb\"/);
  assert.doesNotMatch(script, /find \"\\$\\{bundle_dirs\\[@\\]\\}\" -type f -name \"\\*\\.rpm\"/);
});

function writeFakeTool(binDir, name, contents) {
  const p = path.join(binDir, name);
  writeFileSync(p, contents, "utf8");
  chmodSync(p, 0o755);
}

function writeFakeToolchain(binDir) {
  // dpkg -I is only used for logging.
  writeFakeTool(
    binDir,
    "dpkg",
    `#!/usr/bin/env bash
set -euo pipefail
if [[ "\${1:-}" == "-I" ]]; then
  echo "fake dpkg -I \${2:-}"
  exit 0
fi
echo "fake dpkg: unsupported args: $*" >&2
exit 2
`,
  );

  // dpkg-deb -f/-x are used for metadata and extraction.
  writeFakeTool(
    binDir,
    "dpkg-deb",
    `#!/usr/bin/env bash
set -euo pipefail
cmd="\${1:-}"
case "$cmd" in
  -f)
    field="\${3:-}"
    case "$field" in
      Package) echo "\${FAKE_DEB_PACKAGE:-${expectedMainBinary}}" ;;
      Version) echo "\${FAKE_DEB_VERSION:-${expectedVersion}}" ;;
      Depends) echo "\${FAKE_DEB_DEPENDS:-shared-mime-info, libwebkit2gtk-4.1-0, libgtk-3-0, libayatana-appindicator3-1, librsvg2-2, libssl3}" ;;
      *) echo "" ;;
    esac
    exit 0
    ;;
  -x)
    dest="\${3:-}"
    mkdir -p "$dest/usr/bin" "$dest/usr/share/mime/packages"
    echo "ELF stub" > "$dest/usr/bin/${expectedMainBinary}"
    if [[ "\${FAKE_DEB_WRITE_MIME_XML:-1}" == "1" ]]; then
      if [[ -n "\${FAKE_DEB_MIME_XML_CONTENT:-}" ]]; then
        printf '%s\\n' "$FAKE_DEB_MIME_XML_CONTENT" > "$dest/usr/share/mime/packages/${expectedIdentifier}.xml"
      else
        cat > "$dest/usr/share/mime/packages/${expectedIdentifier}.xml" <<'XML'
<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet">
    <glob pattern="*.parquet" />
  </mime-type>
</mime-info>
XML
      fi
    fi
    exit 0
    ;;
  *)
    echo "fake dpkg-deb: unsupported args: $*" >&2
    exit 2
    ;;
esac
`,
  );

  // rpm -qp --queryformat and rpm -qpR are used for metadata checks.
  writeFakeTool(
    binDir,
    "rpm",
    `#!/usr/bin/env bash
set -euo pipefail
cmd="\${1:-}"
if [[ "$cmd" == "-qpR" ]]; then
  if [[ -z "\${FAKE_RPM_REQUIRES_FILE:-}" ]]; then
    echo "fake rpm: missing FAKE_RPM_REQUIRES_FILE" >&2
    exit 2
  fi
  cat "$FAKE_RPM_REQUIRES_FILE"
  exit 0
fi
if [[ "$cmd" != "-qp" ]]; then
  echo "fake rpm: unsupported args: $*" >&2
  exit 2
fi
query="\${2:-}"
if [[ "$query" != "--queryformat" ]]; then
  echo "fake rpm: unsupported query: $query" >&2
  exit 2
fi
fmt="\${3:-}"
if [[ "$fmt" == *"%{VERSION}"* ]]; then
  echo "\${FAKE_RPM_VERSION:-${expectedVersion}}"
  exit 0
fi
if [[ "$fmt" == *"%{NAME}"* ]]; then
  echo "\${FAKE_RPM_NAME:-${expectedMainBinary}}"
  exit 0
fi
echo "fake rpm: unsupported queryformat: $fmt" >&2
exit 2
`,
  );

  // rpm2cpio/cpio are used for extraction; we fake them by making cpio write the desired files.
  writeFakeTool(
    binDir,
    "rpm2cpio",
    `#!/usr/bin/env bash
set -euo pipefail
cat >/dev/null || true
exit 0
`,
  );

  writeFakeTool(
    binDir,
    "cpio",
    `#!/usr/bin/env bash
set -euo pipefail
cat >/dev/null || true
mkdir -p usr/bin usr/share/mime/packages
echo "ELF stub" > usr/bin/${expectedMainBinary}
if [[ "\${FAKE_RPM_WRITE_MIME_XML:-1}" == "1" ]]; then
  if [[ -n "\${FAKE_RPM_MIME_XML_CONTENT:-}" ]]; then
    printf '%s\\n' "$FAKE_RPM_MIME_XML_CONTENT" > "usr/share/mime/packages/${expectedIdentifier}.xml"
  else
    cat > "usr/share/mime/packages/${expectedIdentifier}.xml" <<'XML'
<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet">
    <glob pattern="*.parquet" />
  </mime-type>
</mime-info>
XML
  fi
fi
exit 0
`,
  );

  // Make the stripped binary checks deterministic.
  writeFakeTool(
    binDir,
    "file",
    `#!/usr/bin/env bash
set -euo pipefail
if [[ "\${1:-}" == "-b" ]]; then
  echo "ELF 64-bit LSB executable, x86-64, stripped"
  exit 0
fi
echo "fake file: unsupported args: $*" >&2
exit 2
`,
  );

  writeFakeTool(
    binDir,
    "readelf",
    `#!/usr/bin/env bash
set -euo pipefail
# The verifier only checks that output does not contain '.debug_'.
echo "Section Headers:"
echo "  [ 0] .text PROGBITS"
exit 0
`,
  );
}

function runVerifier({ cwd, env } = {}) {
  const proc = spawnSync("bash", [path.join(repoRoot, "scripts", "ci", "verify-linux-package-deps.sh")], {
    cwd: cwd ?? repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      ...env,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("verify-linux-package-deps passes when bundles include Parquet shared-mime-info XML with expected content", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const cargoTargetDir = path.join(tmp, "target");
  const debDir = path.join(cargoTargetDir, "release", "bundle", "deb");
  const rpmDir = path.join(cargoTargetDir, "release", "bundle", "rpm");
  mkdirSync(debDir, { recursive: true });
  mkdirSync(rpmDir, { recursive: true });
  writeFileSync(path.join(debDir, "Formula.deb"), "not-a-real-deb", "utf8");
  writeFileSync(path.join(rpmDir, "Formula.rpm"), "not-a-real-rpm", "utf8");

  const requiresFile = path.join(tmp, "rpm-requires.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      CARGO_TARGET_DIR: cargoTargetDir,
      FAKE_RPM_REQUIRES_FILE: requiresFile,
      FAKE_DEB_VERSION: expectedVersion,
      FAKE_DEB_PACKAGE: expectedMainBinary,
      FAKE_RPM_VERSION: expectedVersion,
      FAKE_RPM_NAME: expectedMainBinary,
    },
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify-linux-package-deps guardrails validate identifier-derived MIME XML filename (no path separators)", { skip: !hasBash }, () => {
  const script = stripHashComments(readFileSync(path.join(repoRoot, "scripts", "ci", "verify-linux-package-deps.sh"), "utf8"));
  assert.match(script, /contains path separators/i);
});

test("verify-linux-package-deps fails when tauri identifier contains path separators (Parquet configured)", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const tauriConfPath = path.join(tmp, "tauri.conf.json");
  writeFileSync(
    tauriConfPath,
    JSON.stringify(
      {
        version: expectedVersion,
        mainBinaryName: expectedMainBinary,
        identifier: "com/example.formula.desktop",
        bundle: {
          fileAssociations: [
            {
              ext: ["parquet"],
              mimeType: "application/vnd.apache.parquet",
            },
          ],
        },
      },
      null,
      2,
    ),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      FORMULA_TAURI_CONF_PATH: tauriConfPath,
    },
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /path separators/i);
});

test(
  "verify-linux-package-deps fails when Parquet association is configured but tauri identifier is missing",
  { skip: !hasBash },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
    const binDir = path.join(tmp, "bin");
    mkdirSync(binDir, { recursive: true });
    writeFakeToolchain(binDir);

    const tauriConfPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      tauriConfPath,
      JSON.stringify(
        {
          version: expectedVersion,
          mainBinaryName: expectedMainBinary,
          bundle: {
            fileAssociations: [
              {
                ext: ["parquet"],
                mimeType: "application/vnd.apache.parquet",
              },
            ],
          },
        },
        null,
        2,
      ),
      "utf8",
    );

    const proc = runVerifier({
      env: {
        PATH: `${binDir}:${process.env.PATH}`,
        FORMULA_TAURI_CONF_PATH: tauriConfPath,
      },
    });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identifier.*missing/i);
  },
);

test("verify-linux-package-deps fails when RPM Parquet shared-mime-info XML is missing expected content", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const cargoTargetDir = path.join(tmp, "target");
  const debDir = path.join(cargoTargetDir, "release", "bundle", "deb");
  const rpmDir = path.join(cargoTargetDir, "release", "bundle", "rpm");
  mkdirSync(debDir, { recursive: true });
  mkdirSync(rpmDir, { recursive: true });
  writeFileSync(path.join(debDir, "Formula.deb"), "not-a-real-deb", "utf8");
  writeFileSync(path.join(rpmDir, "Formula.rpm"), "not-a-real-rpm", "utf8");

  const requiresFile = path.join(tmp, "rpm-requires.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      CARGO_TARGET_DIR: cargoTargetDir,
      FAKE_RPM_REQUIRES_FILE: requiresFile,
      // Corrupt RPM MIME XML: omit the '*.parquet' glob.
      FAKE_RPM_MIME_XML_CONTENT: `<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet" />
</mime-info>`,
      FAKE_DEB_VERSION: expectedVersion,
      FAKE_DEB_PACKAGE: expectedMainBinary,
      FAKE_RPM_VERSION: expectedVersion,
      FAKE_RPM_NAME: expectedMainBinary,
    },
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info definition file is missing expected content/i);
});

test("verify-linux-package-deps fails when DEB Parquet shared-mime-info XML is missing expected content", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const cargoTargetDir = path.join(tmp, "target");
  const debDir = path.join(cargoTargetDir, "release", "bundle", "deb");
  const rpmDir = path.join(cargoTargetDir, "release", "bundle", "rpm");
  mkdirSync(debDir, { recursive: true });
  mkdirSync(rpmDir, { recursive: true });
  writeFileSync(path.join(debDir, "Formula.deb"), "not-a-real-deb", "utf8");
  writeFileSync(path.join(rpmDir, "Formula.rpm"), "not-a-real-rpm", "utf8");

  const requiresFile = path.join(tmp, "rpm-requires.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      CARGO_TARGET_DIR: cargoTargetDir,
      FAKE_RPM_REQUIRES_FILE: requiresFile,
      // Corrupt DEB MIME XML: omit the '*.parquet' glob.
      FAKE_DEB_MIME_XML_CONTENT: `<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet" />
</mime-info>`,
      FAKE_DEB_VERSION: expectedVersion,
      FAKE_DEB_PACKAGE: expectedMainBinary,
      FAKE_RPM_VERSION: expectedVersion,
      FAKE_RPM_NAME: expectedMainBinary,
    },
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /shared-mime-info definition file is missing expected content/i);
});

test("verify-linux-package-deps fails when DEB Parquet shared-mime-info XML file is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const cargoTargetDir = path.join(tmp, "target");
  const debDir = path.join(cargoTargetDir, "release", "bundle", "deb");
  const rpmDir = path.join(cargoTargetDir, "release", "bundle", "rpm");
  mkdirSync(debDir, { recursive: true });
  mkdirSync(rpmDir, { recursive: true });
  writeFileSync(path.join(debDir, "Formula.deb"), "not-a-real-deb", "utf8");
  writeFileSync(path.join(rpmDir, "Formula.rpm"), "not-a-real-rpm", "utf8");

  const requiresFile = path.join(tmp, "rpm-requires.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      CARGO_TARGET_DIR: cargoTargetDir,
      FAKE_RPM_REQUIRES_FILE: requiresFile,
      FAKE_DEB_VERSION: expectedVersion,
      FAKE_DEB_PACKAGE: expectedMainBinary,
      // Prevent dpkg-deb -x from writing the MIME XML definition file.
      FAKE_DEB_WRITE_MIME_XML: "0",
      FAKE_RPM_VERSION: expectedVersion,
      FAKE_RPM_NAME: expectedMainBinary,
    },
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing Parquet shared-mime-info definition file/i);
});

test("verify-linux-package-deps fails when RPM Parquet shared-mime-info XML file is missing", { skip: !hasBash }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-verify-linux-package-deps-"));
  const binDir = path.join(tmp, "bin");
  mkdirSync(binDir, { recursive: true });
  writeFakeToolchain(binDir);

  const cargoTargetDir = path.join(tmp, "target");
  const debDir = path.join(cargoTargetDir, "release", "bundle", "deb");
  const rpmDir = path.join(cargoTargetDir, "release", "bundle", "rpm");
  mkdirSync(debDir, { recursive: true });
  mkdirSync(rpmDir, { recursive: true });
  writeFileSync(path.join(debDir, "Formula.deb"), "not-a-real-deb", "utf8");
  writeFileSync(path.join(rpmDir, "Formula.rpm"), "not-a-real-rpm", "utf8");

  const requiresFile = path.join(tmp, "rpm-requires.txt");
  writeFileSync(
    requiresFile,
    [
      "shared-mime-info",
      "(webkit2gtk4.1 or libwebkit2gtk-4_1-0)",
      "(gtk3 or libgtk-3-0)",
      "((libayatana-appindicator-gtk3 or libappindicator-gtk3) or (libayatana-appindicator3-1 or libappindicator3-1))",
      "(librsvg2 or librsvg-2-2)",
      "(openssl-libs or libopenssl3)",
    ].join("\n"),
    "utf8",
  );

  const proc = runVerifier({
    env: {
      PATH: `${binDir}:${process.env.PATH}`,
      CARGO_TARGET_DIR: cargoTargetDir,
      FAKE_RPM_REQUIRES_FILE: requiresFile,
      FAKE_DEB_VERSION: expectedVersion,
      FAKE_DEB_PACKAGE: expectedMainBinary,
      FAKE_RPM_VERSION: expectedVersion,
      FAKE_RPM_NAME: expectedMainBinary,
      // Prevent rpm2cpio/cpio extraction from writing the MIME XML definition file.
      FAKE_RPM_WRITE_MIME_XML: "0",
    },
  });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing Parquet shared-mime-info definition file/i);
});
