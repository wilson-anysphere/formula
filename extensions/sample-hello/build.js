const fs = require("node:fs/promises");
const path = require("node:path");

const SRC_PATH = path.join(__dirname, "src", "extension.js");
const DIST_CJS_PATH = path.join(__dirname, "dist", "extension.js");
const DIST_ESM_PATH = path.join(__dirname, "dist", "extension.mjs");

const GENERATED_HEADER = `// This file is generated from src/extension.js by build.js. Do not edit.\n`;

function toEsm(src) {
  let out = src;

  out = out.replace(
    /^const formula = require\(["']@formula\/extension-api["']\);\s*\n/m,
    'import * as formula from "@formula/extension-api";\n'
  );

  out = out.replace(/\nmodule\.exports\s*=\s*\{\s*activate\s*\};?\s*$/m, "\nexport { activate };\n");

  out = out.replace(/\nmodule\.exports\s*=\s*\{\s*\n\s*activate\s*\n\s*\};?\s*$/m, "\nexport { activate };\n");

  return out;
}

async function build({ check = false } = {}) {
  const src = await fs.readFile(SRC_PATH, "utf8");
  const cjsOutput = GENERATED_HEADER + src;
  const esmOutput = GENERATED_HEADER + toEsm(src);

  if (check) {
    let distCjs;
    let distEsm;
    try {
      distCjs = await fs.readFile(DIST_CJS_PATH, "utf8");
    } catch (error) {
      const message = error && error.code === "ENOENT"
        ? `Missing ${path.relative(process.cwd(), DIST_CJS_PATH)}. Run: node ${path.relative(process.cwd(), path.join(__dirname, "build.js"))}`
        : `Failed to read ${path.relative(process.cwd(), DIST_CJS_PATH)}: ${error?.message ?? error}`;
      throw new Error(message);
    }

    try {
      distEsm = await fs.readFile(DIST_ESM_PATH, "utf8");
    } catch (error) {
      const message = error && error.code === "ENOENT"
        ? `Missing ${path.relative(process.cwd(), DIST_ESM_PATH)}. Run: node ${path.relative(process.cwd(), path.join(__dirname, "build.js"))}`
        : `Failed to read ${path.relative(process.cwd(), DIST_ESM_PATH)}: ${error?.message ?? error}`;
      throw new Error(message);
    }

    if (distCjs !== cjsOutput) {
      throw new Error(
        `extensions/sample-hello/dist/extension.js is out of date. Run: node ${path.relative(
          process.cwd(),
          path.join(__dirname, "build.js")
        )}`
      );
    }

    if (distEsm !== esmOutput) {
      throw new Error(
        `extensions/sample-hello/dist/extension.mjs is out of date. Run: node ${path.relative(
          process.cwd(),
          path.join(__dirname, "build.js")
        )}`
      );
    }
    return;
  }

  await fs.mkdir(path.dirname(DIST_CJS_PATH), { recursive: true });
  await fs.writeFile(DIST_CJS_PATH, cjsOutput, "utf8");
  await fs.writeFile(DIST_ESM_PATH, esmOutput, "utf8");
}

if (require.main === module) {
  const check = process.argv.includes("--check");
  build({ check }).catch((error) => {
    // eslint-disable-next-line no-console
    console.error(error?.message ?? error);
    process.exit(1);
  });
}

module.exports = {
  build,
  SRC_PATH,
  DIST_CJS_PATH,
  DIST_ESM_PATH,
  GENERATED_HEADER,
  toEsm
};
