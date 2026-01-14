import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripHashComments } from "../../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-appimage.sh");

const hasBash = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("bash", ["-lc", "exit 0"], { stdio: "ignore" });
  return probe.status === 0;
})();

const hasPython3 = (() => {
  if (process.platform === "win32") return false;
  const probe = spawnSync("python3", ["-c", "import sys; sys.exit(0)"], { stdio: "ignore" });
  return probe.status === 0;
})();

test("check-appimage: --help prints usage and mentions FORMULA_TAURI_CONF_PATH", { skip: !hasBash }, () => {
  const proc = spawnSync("bash", [scriptPath, "--help"], { cwd: repoRoot, encoding: "utf8" });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /check-appimage\.sh/i);
  assert.match(proc.stdout, /FORMULA_TAURI_CONF_PATH/);
});

test("check-appimage avoids unbounded find scans when directory args are provided (perf guardrail)", () => {
  const raw = stripHashComments(fs.readFileSync(scriptPath, "utf8"));
  // Historical versions used `find "$arg" -type f -name '*.AppImage'` which can be
  // extremely slow when callers pass a Cargo `target/` directory.
  assert.ok(
    !raw.includes(`find "$arg" -type f -name '*.AppImage' -print0`),
    "Expected check-appimage.sh to avoid unbounded `find \"$arg\" -type f -name '*.AppImage'` scans.",
  );
});

test(
  "check-appimage: honors FORMULA_TAURI_CONF_PATH when validating desktop integration",
  { skip: !hasBash || !hasPython3 || process.platform !== "linux" },
  () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "formula-check-appimage-"));
    try {
      const identifier = "com.example.override";
      const mainBinaryName = "override-app";
      const scheme = "override";

      const binDir = path.join(tmp, "bin");
      fs.mkdirSync(binDir, { recursive: true });

      // Stub system commands so we can run the smoke test without producing real ELF binaries.
      fs.writeFileSync(
        path.join(binDir, "file"),
        `#!/usr/bin/env bash
set -euo pipefail
base=0
for arg in "$@"; do
  if [[ "$arg" == "-b" ]]; then
    base=1
  fi
done
while [[ $# -gt 0 && "$1" == -* ]]; do
  shift
done
target="\${1:-}"
arch="$(uname -m)"
case "$arch" in
  x86_64) arch_str="x86-64" ;;
  aarch64) arch_str="aarch64" ;;
  armv7l) arch_str="ARM" ;;
  *) arch_str="$arch" ;;
esac
desc="ELF 64-bit LSB executable, $arch_str, version 1 (SYSV), dynamically linked, stripped"
if [[ "$base" -eq 1 ]]; then
  echo "$desc"
else
  echo "$target: $desc"
fi
`,
        { mode: 0o755 },
      );

      fs.writeFileSync(
        path.join(binDir, "readelf"),
        `#!/usr/bin/env bash
set -euo pipefail
arch="$(uname -m)"
case "$arch" in
  x86_64) machine="Advanced Micro Devices X86-64" ;;
  aarch64) machine="AArch64" ;;
  armv7l) machine="ARM" ;;
  *) machine="$arch" ;;
esac
if [[ " $* " == *" -h "* ]]; then
  cat <<EOF
ELF Header:
  Machine:                           $machine
EOF
  exit 0
fi
if [[ " $* " == *" -S "* ]]; then
  cat <<'EOF'
Section Headers:
  [ 0] .text PROGBITS 0000000000000000 000000 000000 00  AX  0 0 16
EOF
  exit 0
fi
echo "readelf stub: unsupported args: $*" >&2
exit 1
`,
        { mode: 0o755 },
      );

      fs.writeFileSync(
        path.join(binDir, "ldd"),
        `#!/usr/bin/env bash
set -euo pipefail
cat <<'EOF'
linux-vdso.so.1 (0x00007fff00000000)
libc.so.6 => /lib/libc.so.6 (0x00007fff00000000)
/lib64/ld-linux-x86-64.so.2 (0x00007fff00000000)
EOF
`,
        { mode: 0o755 },
      );

      const tauriConfigPath = path.join(tmp, "tauri.override.json");
      fs.writeFileSync(
        tauriConfigPath,
        JSON.stringify(
          {
            mainBinaryName,
            identifier,
            bundle: {
              fileAssociations: [
                {
                  ext: ["parquet"],
                  mimeType: "application/vnd.apache.parquet",
                  name: "Parquet File",
                  role: "Editor",
                },
              ],
            },
            plugins: { "deep-link": { desktop: { schemes: [scheme] } } },
          },
          null,
          2,
        ),
      );

      const appImagePath = path.join(tmp, "Fake.AppImage");
      fs.writeFileSync(
        appImagePath,
        `#!/usr/bin/env bash
set -euo pipefail
if [[ "\${1:-}" == "--appimage-extract" ]]; then
  mkdir -p squashfs-root/usr/bin
  mkdir -p squashfs-root/usr/share/doc/${mainBinaryName}
  mkdir -p squashfs-root/usr/share/mime/packages

  cat > squashfs-root/usr/bin/${mainBinaryName} <<'EOF'
#!/usr/bin/env bash
exit 0
EOF
  chmod +x squashfs-root/usr/bin/${mainBinaryName}

  cat > squashfs-root/${mainBinaryName}.desktop <<'EOF'
[Desktop Entry]
Type=Application
Name=Override App
Exec=${mainBinaryName} %U
MimeType=application/vnd.apache.parquet;x-scheme-handler/${scheme};
EOF

  echo "license" > squashfs-root/usr/share/doc/${mainBinaryName}/LICENSE
  echo "notice" > squashfs-root/usr/share/doc/${mainBinaryName}/NOTICE

  cat > squashfs-root/usr/share/mime/packages/${identifier}.xml <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">
  <mime-type type="application/vnd.apache.parquet">
    <glob pattern="*.parquet"/>
  </mime-type>
</mime-info>
EOF

  exit 0
fi
echo "unsupported args: $*" >&2
exit 1
`,
        { mode: 0o755 },
      );

      const proc = spawnSync("bash", [scriptPath, appImagePath], {
        cwd: repoRoot,
        encoding: "utf8",
        env: {
          ...process.env,
          FORMULA_TAURI_CONF_PATH: tauriConfigPath,
          PATH: `${binDir}:${process.env.PATH}`,
        },
      });
      if (proc.error) throw proc.error;
      assert.equal(proc.status, 0, proc.stderr);
      assert.ok(proc.stdout.includes(identifier), proc.stdout);
      assert.ok(proc.stdout.includes(mainBinaryName), proc.stdout);
      assert.ok(proc.stdout.includes(`x-scheme-handler/${scheme}`), proc.stdout);
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  },
);
