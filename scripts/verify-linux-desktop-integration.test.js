import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripPythonComments } from "../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "ci", "verify_linux_desktop_integration.py");

test("verify_linux_desktop_integration avoids Path.rglob() scans (perf guardrail)", () => {
  const contents = stripPythonComments(readFileSync(scriptPath, "utf8"));
  assert.doesNotMatch(contents, /\.rglob\(/, "Expected verifier to avoid unbounded recursive scans");
});

const hasPython3 = (() => {
  const probe = spawnSync("python3", ["--version"], { stdio: "ignore" });
  return !probe.error && probe.status === 0;
})();

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function writeConfig(dir, { mainBinaryName = "formula-desktop", identifier = "app.formula.desktop" } = {}) {
  const configPath = path.join(dir, "tauri.conf.json");
  const fileAssociations = [
    {
      ext: ["xlsx"],
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  ];
  const conf = {
    identifier,
    mainBinaryName,
    bundle: {
      fileAssociations,
    },
  };
  writeFileSync(configPath, JSON.stringify(conf), "utf8");
  return configPath;
}

function writeConfigWithAssociations(
  dir,
  {
    mainBinaryName = "formula-desktop",
    identifier = "app.formula.desktop",
    fileAssociations = [
      {
        ext: ["xlsx"],
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    ],
  } = {},
) {
  const configPath = path.join(dir, "tauri.conf.json");
  const conf = {
    identifier,
    mainBinaryName,
    bundle: { fileAssociations },
  };
  writeFileSync(configPath, JSON.stringify(conf), "utf8");
  return configPath;
}

function writeParquetMimeDefinition(
  pkgRoot,
  {
    identifier = "app.formula.desktop",
    filename = `${identifier}.xml`,
    xmlContent = [
      '<?xml version="1.0" encoding="UTF-8"?>',
      '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
      '  <mime-type type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">',
      '    <glob pattern="*.xlsx" />',
      "  </mime-type>",
      '  <mime-type type="application/vnd.apache.parquet">',
        '    <glob pattern="*.parquet" />',
      "  </mime-type>",
      "</mime-info>",
    ].join("\n"),
  } = {},
) {
  const mimeDir = path.join(pkgRoot, "usr", "share", "mime", "packages");
  mkdirSync(mimeDir, { recursive: true });
  writeFileSync(path.join(mimeDir, filename), xmlContent, "utf8");
}

function writePackageRoot(dir, { execLine, mimeTypeLine, docPackageName = "formula-desktop" } = {}) {
  const pkgRoot = path.join(dir, "pkg");
  const applicationsDir = path.join(pkgRoot, "usr", "share", "applications");
  const docDir = path.join(pkgRoot, "usr", "share", "doc", docPackageName);
  mkdirSync(applicationsDir, { recursive: true });
  mkdirSync(docDir, { recursive: true });
  writeFileSync(path.join(docDir, "LICENSE"), "stub", "utf8");
  writeFileSync(path.join(docDir, "NOTICE"), "stub", "utf8");

  const desktopPath = path.join(applicationsDir, "formula.desktop");
  const exec = execLine ?? "formula-desktop %U";
  const mime =
    mimeTypeLine ??
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;";
  const desktop = [
    "[Desktop Entry]",
    "Name=Formula",
    `Exec=${exec}`,
    `MimeType=${mime}`,
  ].join("\n");
  writeFileSync(desktopPath, desktop, "utf8");
  return pkgRoot;
}

function runValidator({ packageRoot, configPath, extraArgs = [] }) {
  const proc = spawnSync(
    "python3",
    [scriptPath, "--package-root", packageRoot, "--tauri-config", configPath, ...extraArgs],
    { cwd: repoRoot, encoding: "utf8" },
  );
  if (proc.error) throw proc.error;
  return proc;
}

function runValidatorWithEnv({ packageRoot, configPath, extraArgs = [] }) {
  const proc = spawnSync("python3", [scriptPath, "--package-root", packageRoot, ...extraArgs], {
    cwd: repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      FORMULA_TAURI_CONF_PATH: configPath,
    },
  });
  if (proc.error) throw proc.error;
  return proc;
}

test("verify_linux_desktop_integration passes for a desktop entry targeting the expected binary", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp);

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify_linux_desktop_integration supports FORMULA_TAURI_CONF_PATH env override", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp);

  const proc = runValidatorWithEnv({ packageRoot: pkgRoot, configPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test(
  "verify_linux_desktop_integration validates deep-link schemes from config when desktop config is an array",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify(
        {
          identifier: "app.formula.desktop",
          mainBinaryName: "formula-desktop",
          plugins: {
            "deep-link": {
              desktop: [
                {
                  schemes: ["formula", "formula-extra"],
                },
              ],
            },
          },
          bundle: {
            fileAssociations: [
              {
                ext: ["xlsx"],
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              },
            ],
          },
        },
        null,
        2,
      ),
      "utf8",
    );

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;x-scheme-handler/formula-extra;",
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "verify_linux_desktop_integration normalizes deep-link schemes like formula:// from config",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify(
        {
          identifier: "app.formula.desktop",
          mainBinaryName: "formula-desktop",
          plugins: {
            "deep-link": {
              desktop: {
                schemes: ["formula://", "formula-extra/"],
              },
            },
          },
          bundle: {
            fileAssociations: [
              {
                ext: ["xlsx"],
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              },
            ],
          },
        },
        null,
        2,
      ),
      "utf8",
    );

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;x-scheme-handler/formula-extra;",
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "verify_linux_desktop_integration fails when tauri.conf.json deep-link schemes include an invalid value like formula://evil",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify(
        {
          identifier: "app.formula.desktop",
          mainBinaryName: "formula-desktop",
          plugins: {
            "deep-link": {
              desktop: {
                schemes: ["formula", "formula://evil"],
              },
            },
          },
          bundle: {
            fileAssociations: [
              {
                ext: ["xlsx"],
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              },
            ],
          },
        },
        null,
        2,
      ),
      "utf8",
    );

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;",
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /invalid deep-link scheme/i);
    assert.match(proc.stderr, /formula:\/\//i);
  },
);

test(
  "verify_linux_desktop_integration fails when a configured deep-link scheme is missing from the app .desktop MimeType=",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify(
        {
          identifier: "app.formula.desktop",
          mainBinaryName: "formula-desktop",
          plugins: {
            "deep-link": {
              desktop: [
                {
                  schemes: ["formula", "formula-extra"],
                },
              ],
            },
          },
          bundle: {
            fileAssociations: [
              {
                ext: ["xlsx"],
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              },
            ],
          },
        },
        null,
        2,
      ),
      "utf8",
    );

    // Missing x-scheme-handler/formula-extra.
    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;x-scheme-handler/formula;",
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /missing deep link scheme handler/i);
    assert.match(proc.stderr, /x-scheme-handler\/formula-extra/i);
  },
);

test("verify_linux_desktop_integration accepts quoted Exec= paths", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp, { execLine: '"/usr/bin/formula-desktop" %U' });

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.equal(proc.status, 0, proc.stderr);
});

test("verify_linux_desktop_integration fails when no .desktop entries target the expected binary", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp, { execLine: "something-else %U" });

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /target the expected executable/i);
});

test("verify_linux_desktop_integration fails when Exec= is missing a file/URL placeholder", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  // Targets the expected binary, but omits %U/%u/%F/%f so file associations cannot pass opened file paths.
  const pkgRoot = writePackageRoot(tmp, { execLine: "formula-desktop" });

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /placeholder/i);
  assert.match(proc.stderr, /%u\/%U\/%f\/%F/i);
});

test(
  "verify_linux_desktop_integration supports overriding doc package name without changing Exec binary target",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfig(tmp, { mainBinaryName: "formula-desktop" });
    const pkgRoot = writePackageRoot(tmp, { docPackageName: "formula-desktop-alt", execLine: "formula-desktop %U" });

    const proc = runValidator({
      packageRoot: pkgRoot,
      configPath,
      extraArgs: ["--doc-package-name", "formula-desktop-alt", "--expected-main-binary", "formula-desktop"],
    });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test("verify_linux_desktop_integration fails when LICENSE is missing from /usr/share/doc/<pkg>/", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp);

  // Remove the LICENSE file after creating the package root.
  // (Avoid adding extra helper plumbing to keep the test focused.)
  const licensePath = path.join(pkgRoot, "usr", "share", "doc", "formula-desktop", "LICENSE");
  try {
    unlinkSync(licensePath);
  } catch {
    // ignore
  }

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing LICENSE\/NOTICE compliance artifacts/i);
  assert.match(proc.stderr, /LICENSE/i);
});

test("verify_linux_desktop_integration fails when NOTICE is missing from /usr/share/doc/<pkg>/", { skip: !hasPython3 }, () => {
  const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
  const configPath = writeConfig(tmp);
  const pkgRoot = writePackageRoot(tmp);

  const noticePath = path.join(pkgRoot, "usr", "share", "doc", "formula-desktop", "NOTICE");
  try {
    unlinkSync(noticePath);
  } catch {
    // ignore
  }

  const proc = runValidator({ packageRoot: pkgRoot, configPath });
  assert.notEqual(proc.status, 0, "expected non-zero exit status");
  assert.match(proc.stderr, /missing LICENSE\/NOTICE compliance artifacts/i);
  assert.match(proc.stderr, /NOTICE/i);
});

test(
  "verify_linux_desktop_integration passes when Parquet association is configured and shared-mime-info XML is packaged",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfigWithAssociations(tmp, {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
    writeParquetMimeDefinition(pkgRoot);

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.equal(proc.status, 0, proc.stderr);
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but the identifier-based MIME XML file is missing",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const identifier = "com.example.formula.desktop";
    const configPath = writeConfigWithAssociations(tmp, {
      identifier,
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
    // Write a Parquet MIME definition, but under a different identifier filename than the config expects.
    writeParquetMimeDefinition(pkgRoot, { identifier: "app.formula.desktop" });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /expected shared-mime-info definition file is missing/i);
    assert.match(proc.stderr, new RegExp(`${escapeRegExp(identifier)}\\.xml`, "i"));
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but shared-mime-info packages dir is missing",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfigWithAssociations(tmp, {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });

    // Ensure the .desktop file advertises Parquet so the verifier reaches the MIME XML check.
    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /shared-mime-info packages dir/i);
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but tauri identifier contains a path separator",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const identifier = "com/example.formula.desktop";
    const configPath = writeConfigWithAssociations(tmp, {
      identifier,
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });
 
    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
 
    // Ensure the shared-mime-info packages dir exists so the verifier reaches the identifier validation.
    mkdirSync(path.join(pkgRoot, "usr", "share", "mime", "packages"), { recursive: true });
 
    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identifier.*not a valid filename/i);
    assert.match(proc.stderr, /path separators/i);
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but MIME XML lacks the *.parquet glob",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfigWithAssociations(tmp, {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
    writeParquetMimeDefinition(pkgRoot, {
      xmlContent: [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
        '  <mime-type type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">',
        '    <glob pattern="*.xlsx" />',
        "  </mime-type>",
        '  <mime-type type="application/vnd.apache.parquet">',
        // Intentionally omit the *.parquet glob mapping.
        "  </mime-type>",
        "</mime-info>",
      ].join("\n"),
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /glob \*\.parquet/i);
  },
);

test(
  "verify_linux_desktop_integration fails when shared-mime-info XML is missing a glob mapping for a configured non-Parquet extension",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfigWithAssociations(tmp, {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
    // Define only Parquet, omitting the xlsx glob mapping.
    writeParquetMimeDefinition(pkgRoot, {
      xmlContent: [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<mime-info xmlns="http://www.freedesktop.org/standards/shared-mime-info">',
        '  <mime-type type="application/vnd.apache.parquet">',
        '    <glob pattern="*.parquet" />',
        "  </mime-type>",
        "</mime-info>",
      ].join("\n"),
    });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /missing required glob mappings/i);
    assert.match(proc.stderr, /xlsx/i);
    assert.match(proc.stderr, /\*\.xlsx/i);
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but MIME XML filename does not match tauri identifier",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = writeConfigWithAssociations(tmp, {
      fileAssociations: [
        {
          ext: ["xlsx"],
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        {
          ext: ["parquet"],
          mimeType: "application/vnd.apache.parquet",
        },
      ],
    });
    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });

    // Package a valid Parquet shared-mime-info definition, but under the wrong filename.
    // The verifier should enforce the identifier-derived filename: <identifier>.xml.
    writeParquetMimeDefinition(pkgRoot, { filename: "other.xml" });

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /\/usr\/share\/mime\/packages\/app\.formula\.desktop\.xml/i);
    assert.match(proc.stderr, /other\.xml/i);
  },
);

test(
  "verify_linux_desktop_integration fails when Parquet association is configured but tauri identifier is missing",
  { skip: !hasPython3 },
  () => {
    const tmp = mkdtempSync(path.join(tmpdir(), "formula-linux-desktop-integration-"));
    const configPath = path.join(tmp, "tauri.conf.json");
    writeFileSync(
      configPath,
      JSON.stringify(
        {
          // Intentionally omit `identifier`.
          mainBinaryName: "formula-desktop",
          bundle: {
            fileAssociations: [
              {
                ext: ["xlsx"],
                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              },
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

    const pkgRoot = writePackageRoot(tmp, {
      mimeTypeLine:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.apache.parquet;x-scheme-handler/formula;",
    });
    // Ensure the shared-mime-info packages dir exists so the verifier reaches the identifier check.
    writeParquetMimeDefinition(pkgRoot);

    const proc = runValidator({ packageRoot: pkgRoot, configPath });
    assert.notEqual(proc.status, 0, "expected non-zero exit status");
    assert.match(proc.stderr, /identifier is missing/i);
  },
);
