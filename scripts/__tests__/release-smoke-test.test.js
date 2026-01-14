import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const smokeTestPath = path.join(repoRoot, "scripts", "release-smoke-test.mjs");

function currentDesktopTag() {
  const tauriConfPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const config = JSON.parse(fs.readFileSync(tauriConfPath, "utf8"));
  const version = typeof config?.version === "string" ? config.version.trim() : "";
  assert.ok(version, "Expected tauri.conf.json to contain a non-empty version");
  return version.startsWith("v") ? version : `v${version}`;
}

test("release-smoke-test: --help prints usage and exits 0", () => {
  const child = spawnSync(process.execPath, [smokeTestPath, "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });
  assert.equal(child.status, 0, `expected exit 0, got ${child.status}\n${child.stderr}`);
  assert.match(child.stdout, /Release smoke test/i);
});

test("release-smoke-test: runs required steps and can forward --help to verifier", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--", "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Check desktop URL scheme \+ file associations/i);
  assert.match(child.stdout, /Check desktop compliance artifacts/i);
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: defaults --tag from GITHUB_REF_NAME", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, "--repo", "owner/repo", "--", "--help"], {
    cwd: repoRoot,
    env: { ...process.env, GITHUB_REF_NAME: tag },
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: defaults --repo from GITHUB_REPOSITORY", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, "--tag", tag, "--", "--help"], {
    cwd: repoRoot,
    env: { ...process.env, GITHUB_REPOSITORY: "owner/repo" },
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: defaults --repo from git remote origin (when GITHUB_REPOSITORY is unset)", () => {
  const tag = currentDesktopTag();
  const env = { ...process.env };
  delete env.GITHUB_REPOSITORY;

  const child = spawnSync(process.execPath, [smokeTestPath, "--tag", tag, "--", "--help"], {
    cwd: repoRoot,
    env,
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: supports --tag= and --repo= forms", () => {
  const tag = currentDesktopTag();
  const child = spawnSync(process.execPath, [smokeTestPath, `--tag=${tag}`, "--repo=owner/repo", "--", "--help"], {
    cwd: repoRoot,
    encoding: "utf8",
  });

  assert.equal(
    child.status,
    0,
    `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.match(child.stdout, /Release smoke test PASSED/i);
});

test("release-smoke-test: --local-bundles skips validators when bundle dirs exist but no artifacts", () => {
  const tag = currentDesktopTag();
  // This test relies on there being no existing Tauri bundle output directories
  // under the standard search roots. On developer machines (or some CI caching
  // setups) these may exist, and we don't want to delete/modify them.
  const hasExistingBundleDirs = [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  ].some((root) => {
    try {
      return (
        fs.existsSync(path.join(root, "release", "bundle")) ||
        fs.readdirSync(root, { withFileTypes: true })
          .filter((d) => d.isDirectory())
          .some((d) => fs.existsSync(path.join(root, d.name, "release", "bundle")))
      );
    } catch {
      return false;
    }
  });

  if (hasExistingBundleDirs) {
    return;
  }

  // Use the OS temp directory (instead of repoRoot/target) so tests that manipulate
  // temp bundle directories don't race with other suites that run `cargo clean` in
  // parallel (which recursively scans `target/`).
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), `formula-release-smoke-test-empty-${process.pid}-`));
  const bundleDir = path.join(tmpRoot, "release", "bundle");
  fs.mkdirSync(bundleDir, { recursive: true });

  try {
    const child = spawnSync(
      process.execPath,
      [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--local-bundles", "--", "--help"],
      { cwd: repoRoot, encoding: "utf8" },
    );

    assert.equal(
      child.status,
      0,
      `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
    );
    assert.match(child.stdout, /Release smoke test PASSED/i);
    assert.match(child.stdout, /\[SKIP\]/);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});

test("release-smoke-test: --local-bundles runs validate-linux-deb.sh with --no-container when docker is unavailable", () => {
  if (process.platform !== "linux") return;

  const tag = currentDesktopTag();
  const tauriConfPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
  const tauriConf = JSON.parse(fs.readFileSync(tauriConfPath, "utf8"));
  const expectedVersion = String(tauriConf?.version ?? "").trim();
  const expectedDebName = String(tauriConf?.mainBinaryName ?? "").trim() || "formula-desktop";
  const expectedIdentifier = String(tauriConf?.identifier ?? "").trim() || "app.formula.desktop";
  const expectedMimeDefinitionContentsPath = `./usr/share/mime/packages/${expectedIdentifier}.xml`;
  assert.ok(expectedVersion, "Expected tauri.conf.json to contain a non-empty version");

  // Like the empty-artifacts test above, avoid relying on / mutating any real bundle outputs
  // that may exist on developer machines.
  const hasExistingBundleDirs = [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  ].some((root) => {
    try {
      return (
        fs.existsSync(path.join(root, "release", "bundle")) ||
        fs.readdirSync(root, { withFileTypes: true })
          .filter((d) => d.isDirectory())
          .some((d) => fs.existsSync(path.join(root, d.name, "release", "bundle")))
      );
    } catch {
      return false;
    }
  });

  if (hasExistingBundleDirs) {
    return;
  }

  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), `formula-release-smoke-test-deb-nodocker-${process.pid}-`));
  const bundleDir = path.join(tmpRoot, "release", "bundle", "deb");
  const binDir = path.join(tmpRoot, "bin");
  fs.mkdirSync(bundleDir, { recursive: true });
  fs.mkdirSync(binDir, { recursive: true });

  const debPath = path.join(bundleDir, "Formula.deb");
  fs.writeFileSync(debPath, "not-a-real-deb", { encoding: "utf8" });

  const deepLinkDesktop = tauriConf?.plugins?.["deep-link"]?.desktop;
  const deepLinkSchemes = new Set();
  const addSchemes = (protocol) => {
    const raw = protocol?.schemes;
    const values = typeof raw === "string" ? [raw] : Array.isArray(raw) ? raw : [];
    for (const v of values) {
      if (typeof v !== "string") continue;
      const normalized = v.trim().replace(/[:/]+$/, "").toLowerCase();
      if (normalized) deepLinkSchemes.add(normalized);
    }
  };
  if (Array.isArray(deepLinkDesktop)) {
    for (const protocol of deepLinkDesktop) addSchemes(protocol);
  } else if (deepLinkDesktop != null) {
    addSchemes(deepLinkDesktop);
  }
  if (deepLinkSchemes.size === 0) deepLinkSchemes.add("formula");
  const schemeMimes = Array.from(deepLinkSchemes)
    .sort()
    .map((scheme) => `x-scheme-handler/${scheme}`);

  const mimeList = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    "application/vnd.ms-excel.template.macroEnabled.12",
    "application/vnd.ms-excel.addin.macroEnabled.12",
    "text/csv",
    "application/vnd.apache.parquet",
    ...schemeMimes,
  ].join(";");

  const dpkgDebStub = `#!/usr/bin/env bash
set -euo pipefail

expected_version="${expectedVersion}"
expected_pkg="${expectedDebName}"

cmd="\${1:-}"
case "$cmd" in
  --version)
    echo "dpkg-deb (stub)"
    exit 0
    ;;
  -f)
    field="\${3:-}"
    case "$field" in
      Version) echo "$expected_version" ;;
      Package) echo "$expected_pkg" ;;
      Depends) echo "shared-mime-info, libwebkit2gtk-4.1-0, libgtk-3-0, libayatana-appindicator3-1, librsvg2-2, libssl3" ;;
      *) echo "" ;;
    esac
    exit 0
    ;;
  -c|--contents)
    cat <<EOF
-rwxr-xr-x root/root 0 2024-01-01 00:00 ./usr/bin/$expected_pkg
-rw-r--r-- root/root 0 2024-01-01 00:00 ./usr/share/applications/formula.desktop
-rw-r--r-- root/root 0 2024-01-01 00:00 ./usr/share/doc/$expected_pkg/LICENSE
-rw-r--r-- root/root 0 2024-01-01 00:00 ./usr/share/doc/$expected_pkg/NOTICE
-rw-r--r-- root/root 0 2024-01-01 00:00 ${expectedMimeDefinitionContentsPath}
EOF
    exit 0
    ;;
  -x)
    dest="\${3:-}"
    mkdir -p "$dest/usr/bin" "$dest/usr/share/applications" "$dest/usr/share/doc/$expected_pkg" "$dest/usr/share/mime/packages"
    cat > "$dest/usr/bin/$expected_pkg" <<'BIN'
#!/usr/bin/env bash
echo "formula stub"
BIN
    chmod +x "$dest/usr/bin/$expected_pkg"
    cat > "$dest/usr/share/applications/formula.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Formula
Exec=$expected_pkg %U
MimeType=${mimeList};
EOF
    echo "LICENSE stub" > "$dest/usr/share/doc/$expected_pkg/LICENSE"
    echo "NOTICE stub" > "$dest/usr/share/doc/$expected_pkg/NOTICE"
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
    echo "dpkg-deb stub: unsupported args: $*" >&2
    exit 2
    ;;
esac
`;
  const dpkgDebPath = path.join(binDir, "dpkg-deb");
  fs.writeFileSync(dpkgDebPath, dpkgDebStub, { encoding: "utf8" });
  fs.chmodSync(dpkgDebPath, 0o755);

  // Stub docker so `docker info` fails, which should force release-smoke-test to pass
  // --no-container to validate-linux-deb.sh.
  const dockerPath = path.join(binDir, "docker");
  fs.writeFileSync(dockerPath, "#!/usr/bin/env bash\nexit 1\n", { encoding: "utf8" });
  fs.chmodSync(dockerPath, 0o755);

  try {
    const child = spawnSync(
      process.execPath,
      [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--local-bundles", "--", "--help"],
      {
        cwd: repoRoot,
        env: {
          ...process.env,
          CARGO_TARGET_DIR: tmpRoot,
          PATH: `${binDir}:${process.env.PATH}`,
        },
        encoding: "utf8",
      },
    );

    assert.equal(
      child.status,
      0,
      `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
    );
    assert.match(child.stdout, /Release smoke test PASSED/i);
    assert.match(child.stdout, /=== Validate local bundles \(linux\): validate-linux-deb\.sh ===/i);
    assert.match(child.stdout, /validate-linux-deb\.sh: OK/i);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});

test("release-smoke-test: --local-bundles skips validate-linux-appimage.sh when unsquashfs is missing", () => {
  if (process.platform !== "linux") return;

  const tag = currentDesktopTag();

  const hasExistingBundleDirs = [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  ].some((root) => {
    try {
      return (
        fs.existsSync(path.join(root, "release", "bundle")) ||
        fs.readdirSync(root, { withFileTypes: true })
          .filter((d) => d.isDirectory())
          .some((d) => fs.existsSync(path.join(root, d.name, "release", "bundle")))
      );
    } catch {
      return false;
    }
  });

  if (hasExistingBundleDirs) {
    return;
  }

  const tmpRoot = fs.mkdtempSync(
    path.join(os.tmpdir(), `formula-release-smoke-test-appimage-nosquashfs-${process.pid}-`),
  );
  const bundleDir = path.join(tmpRoot, "release", "bundle", "appimage");
  const binDir = path.join(tmpRoot, "bin");
  fs.mkdirSync(bundleDir, { recursive: true });
  fs.mkdirSync(binDir, { recursive: true });
  fs.writeFileSync(path.join(bundleDir, "Formula.AppImage"), "not-a-real-appimage", { encoding: "utf8" });

  try {
    const child = spawnSync(
      process.execPath,
      [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--local-bundles", "--", "--help"],
      {
        cwd: repoRoot,
        env: {
          ...process.env,
          CARGO_TARGET_DIR: tmpRoot,
          // Provide a PATH that does not include system binaries so unsquashfs is guaranteed missing,
          // regardless of the host environment.
          PATH: binDir,
        },
        encoding: "utf8",
      },
    );

    assert.equal(
      child.status,
      0,
      `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
    );
    assert.match(child.stdout, /Release smoke test PASSED/i);
    assert.match(child.stdout, /Skipping validate-linux-appimage\.sh.*unsquashfs/i);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});

test("release-smoke-test: --local-bundles skips validate-linux-rpm.sh when docker is unavailable and rpm2cpio/cpio are missing", () => {
  if (process.platform !== "linux") return;

  const tag = currentDesktopTag();

  const hasExistingBundleDirs = [
    path.join(repoRoot, "apps", "desktop", "src-tauri", "target"),
    path.join(repoRoot, "apps", "desktop", "target"),
    path.join(repoRoot, "target"),
  ].some((root) => {
    try {
      return (
        fs.existsSync(path.join(root, "release", "bundle")) ||
        fs.readdirSync(root, { withFileTypes: true })
          .filter((d) => d.isDirectory())
          .some((d) => fs.existsSync(path.join(root, d.name, "release", "bundle")))
      );
    } catch {
      return false;
    }
  });

  if (hasExistingBundleDirs) {
    return;
  }

  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), `formula-release-smoke-test-rpm-nodocker-${process.pid}-`));
  const bundleDir = path.join(tmpRoot, "release", "bundle", "rpm");
  const binDir = path.join(tmpRoot, "bin");
  fs.mkdirSync(bundleDir, { recursive: true });
  fs.mkdirSync(binDir, { recursive: true });
  fs.writeFileSync(path.join(bundleDir, "Formula.rpm"), "not-a-real-rpm", { encoding: "utf8" });

  // Stub rpm so the smoke test believes RPM validation is possible.
  const rpmPath = path.join(binDir, "rpm");
  fs.writeFileSync(rpmPath, "#!/usr/bin/env bash\nexit 0\n", { encoding: "utf8" });
  fs.chmodSync(rpmPath, 0o755);

  // Stub docker so `docker info` fails, forcing the smoke test down the --no-container path.
  const dockerPath = path.join(binDir, "docker");
  fs.writeFileSync(dockerPath, "#!/usr/bin/env bash\nexit 1\n", { encoding: "utf8" });
  fs.chmodSync(dockerPath, 0o755);

  try {
    const child = spawnSync(
      process.execPath,
      [smokeTestPath, "--tag", tag, "--repo", "owner/repo", "--local-bundles", "--", "--help"],
      {
        cwd: repoRoot,
        env: {
          ...process.env,
          CARGO_TARGET_DIR: tmpRoot,
          // Provide a PATH that contains rpm+docker stubs but not rpm2cpio/cpio to force the skip.
          PATH: binDir,
        },
        encoding: "utf8",
      },
    );

    assert.equal(
      child.status,
      0,
      `expected exit 0, got ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
    );
    assert.match(child.stdout, /Release smoke test PASSED/i);
    assert.match(child.stdout, /Skipping validate-linux-rpm\.sh.*rpm2cpio.*cpio/i);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});
