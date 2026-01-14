import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-desktop-compliance-artifacts.mjs");

function run(config) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-desktop-compliance-"));
  const confPath = path.join(tmpdir, "tauri.conf.json");
  writeFileSync(confPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");

  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: confPath,
    },
  });

  rmSync(tmpdir, { recursive: true, force: true });
  if (proc.error) throw proc.error;
  return proc;
}

test("passes when LICENSE/NOTICE resources + Linux doc files are configured", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["../../../LICENSE", "../../../NOTICE"],
      linux: {
        deb: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
        rpm: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
      },
    },
  });
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(proc.stdout, /desktop-compliance: OK/i);
});

test("fails when bundle.resources is missing NOTICE", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["../../../LICENSE"],
      linux: {
        deb: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
        rpm: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bundle\.resources.*NOTICE/i);
});

test("fails when Linux doc dir does not match mainBinaryName", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["../../../LICENSE", "../../../NOTICE"],
      linux: {
        deb: {
          files: {
            "usr/share/doc/other/LICENSE": "../../../LICENSE",
            "usr/share/doc/other/NOTICE": "../../../NOTICE",
          },
        },
        rpm: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "../../../LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "../../../NOTICE",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bundle\.linux\.deb\.files/i);
  assert.match(proc.stderr, /usr\/share\/doc\/formula-desktop\/LICENSE/i);
});

