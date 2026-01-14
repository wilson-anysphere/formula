import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "check-desktop-compliance-artifacts.mjs");

const testIdentifier = "com.example.formula.desktop";

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function run(config, { writeLicense = true, writeNotice = true, mimeXmlContent } = {}) {
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
  const identifier =
    typeof config?.identifier === "string" && config.identifier.trim() ? config.identifier.trim() : "app.formula.desktop";
  const mimeBasename = `${identifier}.xml`;
  const defaultMimeXml = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
    '  <mime-type type="application/vnd.apache.parquet">',
    '    <glob pattern="*.parquet" />',
    "  </mime-type>",
    "</mime-info>",
    "",
  ].join("\n");
  const mimeXml = typeof mimeXmlContent === "string" ? mimeXmlContent : defaultMimeXml;
  mkdirSync(path.join(tmpdir, "mime"), { recursive: true });
  {
    const mimePath = path.join(tmpdir, "mime", mimeBasename);
    mkdirSync(path.dirname(mimePath), { recursive: true });
    writeFileSync(mimePath, mimeXml, "utf8");
  }
  // Some configs may reference the MIME definition file at repo root (basename-only); create
  // a stub for that path too.
  {
    const rootMimePath = path.join(tmpdir, mimeBasename);
    mkdirSync(path.dirname(rootMimePath), { recursive: true });
    writeFileSync(rootMimePath, mimeXml, "utf8");
  }
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

test("honors FORMULA_TAURI_CONF_PATH as a repo-root-relative path (independent of cwd)", () => {
  const tmpRoot = path.join(repoRoot, "target");
  mkdirSync(tmpRoot, { recursive: true });
  const tmpdir = mkdtempSync(path.join(tmpRoot, "desktop-compliance-rel-"));
  try {
    const confPath = path.join(tmpdir, "tauri.conf.json");
    const repoLicense = path.join(repoRoot, "LICENSE");
    const repoNotice = path.join(repoRoot, "NOTICE");
    writeFileSync(
      confPath,
      `${JSON.stringify(
        {
          mainBinaryName: "formula-desktop",
          bundle: {
            resources: [repoLicense, repoNotice],
            fileAssociations: [],
            linux: {
              deb: {
                files: {
                  "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                  "usr/share/doc/formula-desktop/NOTICE": repoNotice,
                },
              },
              rpm: {
                files: {
                  "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                  "usr/share/doc/formula-desktop/NOTICE": repoNotice,
                },
              },
              appimage: {
                files: {
                  "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                  "usr/share/doc/formula-desktop/NOTICE": repoNotice,
                },
              },
            },
          },
        },
        null,
        2,
      )}\n`,
      "utf8",
    );

    const proc = spawnSync(process.execPath, [scriptPath], {
      cwd: path.join(repoRoot, "apps", "desktop"),
      encoding: "utf8",
      env: {
        ...process.env,
        FORMULA_TAURI_CONF_PATH: path.relative(repoRoot, confPath),
      },
    });
    if (proc.error) throw proc.error;
    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
    assert.match(proc.stdout, /desktop-compliance: OK/i);
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
});

test("fails when tauri.conf.json (in repo) points LICENSE at a non-root LICENSE file via bundle.linux.*.files", () => {
  const tmpRoot = path.join(repoRoot, "target");
  mkdirSync(tmpRoot, { recursive: true });
  const tmpdir = mkdtempSync(path.join(tmpRoot, "desktop-compliance-inrepo-"));
  const confPath = path.join(tmpdir, "tauri.conf.json");

  // Use absolute paths so the config can live anywhere under repoRoot.
  const repoLicense = path.join(repoRoot, "LICENSE");
  const repoNotice = path.join(repoRoot, "NOTICE");
  const otherLicense = path.join(repoRoot, "crates", "formula-wasm", "LICENSE");

  writeFileSync(
    confPath,
    `${JSON.stringify(
      {
        mainBinaryName: "formula-desktop",
        bundle: {
          resources: [repoLicense, repoNotice],
          linux: {
            deb: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": otherLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
            rpm: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": otherLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
            appimage: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": otherLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
          },
        },
      },
      null,
      2,
    )}\n`,
    "utf8",
  );

  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env, FORMULA_TAURI_CONF_PATH: confPath },
  });

  rmSync(tmpdir, { recursive: true, force: true });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bundle\.linux\.deb\.files/i);
  assert.match(proc.stderr, /repo root LICENSE/i);
});

test("fails when tauri.conf.json (in repo) points LICENSE at a non-root LICENSE file via bundle.resources", () => {
  const tmpRoot = path.join(repoRoot, "target");
  mkdirSync(tmpRoot, { recursive: true });
  const tmpdir = mkdtempSync(path.join(tmpRoot, "desktop-compliance-inrepo-"));
  const confPath = path.join(tmpdir, "tauri.conf.json");

  const repoLicense = path.join(repoRoot, "LICENSE");
  const repoNotice = path.join(repoRoot, "NOTICE");
  const otherLicense = path.join(repoRoot, "crates", "formula-wasm", "LICENSE");

  writeFileSync(
    confPath,
    `${JSON.stringify(
      {
        mainBinaryName: "formula-desktop",
        bundle: {
          resources: [otherLicense, repoNotice],
          linux: {
            deb: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
            rpm: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
            appimage: {
              files: {
                "usr/share/doc/formula-desktop/LICENSE": repoLicense,
                "usr/share/doc/formula-desktop/NOTICE": repoNotice,
              },
            },
          },
        },
      },
      null,
      2,
    )}\n`,
    "utf8",
  );

  const proc = spawnSync(process.execPath, [scriptPath], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env, FORMULA_TAURI_CONF_PATH: confPath },
  });

  rmSync(tmpdir, { recursive: true, force: true });

  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /bundle\.resources/i);
  assert.match(proc.stderr, /repo root LICENSE/i);
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
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const mimeSrc = `mime/${testIdentifier}.xml`;
  const proc = run({
    identifier: testIdentifier,
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
            [mimeDest]: mimeSrc,
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            [mimeDest]: mimeSrc,
          },
        },
        appimage: {
          files: {
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            [mimeDest]: mimeSrc,
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
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const mimeSrc = `mime/${testIdentifier}.xml`;
  const proc = run({
    identifier: testIdentifier,
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["shared-mime-info"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
      },
    },
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when Parquet association is configured but identifier contains a path separator", () => {
  const badIdentifier = "com/example.formula.desktop";
  const mimeDest = `usr/share/mime/packages/${badIdentifier}.xml`;
  const mimeSrc = `mime/${badIdentifier}.xml`;
  const proc = run({
    identifier: badIdentifier,
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["shared-mime-info"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["shared-mime-info"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
      },
    },
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /identifier must be a valid filename/i);
  assert.match(proc.stderr, /path separators/i);
});

test("fails when Parquet association is configured but Linux bundles omit MIME definition mapping", () => {
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const proc = run({
    identifier: testIdentifier,
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
  assert.match(proc.stderr, new RegExp(escapeRegExp(mimeDest), "i"));
});

test("fails when Parquet association is configured but Linux package deps omit shared-mime-info", () => {
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const mimeSrc = `mime/${testIdentifier}.xml`;
  const proc = run({
    identifier: testIdentifier,
    mainBinaryName: "formula-desktop",
    bundle: {
      resources: ["LICENSE", "NOTICE"],
      fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
      linux: {
        deb: {
          depends: ["libwebkit2gtk-4.1-0"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        rpm: {
          depends: ["(webkit2gtk4.1 or libwebkit2gtk-4_1-0)"],
          files: {
            [mimeDest]: mimeSrc,
            "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
            "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
          },
        },
        appimage: {
          files: {
            [mimeDest]: mimeSrc,
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

test("fails when Parquet shared-mime-info definition file lacks expected content", () => {
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const mimeSrc = `mime/${testIdentifier}.xml`;
  const proc = run(
    {
      identifier: testIdentifier,
      mainBinaryName: "formula-desktop",
      bundle: {
        resources: ["LICENSE", "NOTICE"],
        fileAssociations: [{ ext: ["parquet"], mimeType: "application/vnd.apache.parquet" }],
        linux: {
          deb: {
            depends: ["shared-mime-info"],
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          rpm: {
            depends: ["shared-mime-info"],
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          appimage: {
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
        },
      },
    },
    { mimeXmlContent: "<mime-info />\n" },
  );
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /shared-mime-info definition file is missing required glob mappings/i);
  assert.match(proc.stderr, /\.parquet/i);
});

test("fails when shared-mime-info XML is missing a glob mapping for a configured non-Parquet extension", () => {
  const mimeDest = `usr/share/mime/packages/${testIdentifier}.xml`;
  const mimeSrc = `mime/${testIdentifier}.xml`;
  const proc = run(
    {
      identifier: testIdentifier,
      mainBinaryName: "formula-desktop",
      bundle: {
        resources: ["LICENSE", "NOTICE"],
        fileAssociations: [
          { ext: ["xlsx"], mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
          { ext: ["parquet"], mimeType: "application/vnd.apache.parquet" },
        ],
        linux: {
          deb: {
            depends: ["shared-mime-info"],
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          rpm: {
            depends: ["shared-mime-info"],
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
          appimage: {
            files: {
              [mimeDest]: mimeSrc,
              "usr/share/doc/formula-desktop/LICENSE": "LICENSE",
              "usr/share/doc/formula-desktop/NOTICE": "NOTICE",
            },
          },
        },
      },
    },
    {
      // Only define Parquet mapping; xlsx is configured but missing in the XML.
      mimeXmlContent: [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
        '  <mime-type type="application/vnd.apache.parquet">',
        '    <glob pattern="*.parquet" />',
        "  </mime-type>",
        "</mime-info>",
        "",
      ].join("\n"),
    },
  );
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /missing required glob mappings/i);
  assert.match(proc.stderr, /xlsx/i);
  assert.match(proc.stderr, /\*\.xlsx/i);
});
