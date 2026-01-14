import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-desktop-compliance-artifacts.mjs");

function run(config, { writeLicense = true, writeNotice = true } = {}) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-desktop-compliance-"));
  const confPath = path.join(tmpdir, "tauri.conf.json");
  // The validator resolves paths relative to the tauri.conf.json location. Create
  // dummy LICENSE/NOTICE files next to the temp config so existence checks pass.
  if (writeLicense) {
    writeFileSync(path.join(tmpdir, "LICENSE"), "LICENSE stub\n", "utf8");
  }
  if (writeNotice) {
    writeFileSync(path.join(tmpdir, "NOTICE"), "NOTICE stub\n", "utf8");
  }
  mkdirSync(path.join(tmpdir, "mime"), { recursive: true });
  writeFileSync(path.join(tmpdir, "mime", "app.formula.desktop.xml"), "<mime-info />\n", "utf8");
  // Some configs may reference the MIME definition file at repo root (basename-only); create
  // a stub for that path too.
  writeFileSync(path.join(tmpdir, "app.formula.desktop.xml"), "<mime-info />\n", "utf8");
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
      resources: ["LICENSE", "NOTICE"],
      linux: {
        deb: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
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
      resources: ["LICENSE"],
      linux: {
        deb: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
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
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/doc/other/LICENSE": "LICENSE",
            "usr/share/doc/other/NOTICE": "NOTICE",
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bundle\.linux\.deb\.files/i);
  assert.match(proc.stderr, /usr\/share\/doc\/formula-desktop\/LICENSE/i);
});

test("passes when Parquet MIME file + shared-mime-info deps are configured", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
      },
    },
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when Parquet association is configured but Linux bundles omit MIME definition mapping", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /usr\/share\/mime\/packages\/app\.formula\.desktop\.xml/i);
});

test("fails when Parquet association is configured but Linux package deps omit shared-mime-info", () => {
  const proc = run({
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["libwebkit2gtk-4.1-0"],
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["(webkit2gtk4.1 or libwebkit2gtk-4_1-0)"],
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            "usr/share/mime/packages/app.formula.desktop.xml": "mime/app.formula.desktop.xml",
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /shared-mime-info/i);
});

test("fails when LICENSE source file does not exist", () => {
  const proc = run(
    {
      mainBinaryName: "formula-desktop",
      bundle: {
        resources: ["LICENSE", "NOTICE"],
        linux: {
          deb: {
            files: {
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          rpm: {
            files: {
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          appimage: {
            files: {
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
        },
      },
    },
    { writeLicense: false, writeNotice: true },
  );
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing source file/i);
  assert.match(proc.stderr, /LICENSE/i);
});
