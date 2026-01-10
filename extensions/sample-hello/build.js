const fs = require("node:fs/promises");
const path = require("node:path");

const SRC_PATH = path.join(__dirname, "src", "extension.js");
const DIST_PATH = path.join(__dirname, "dist", "extension.js");

const GENERATED_HEADER = `// This file is generated from src/extension.js by build.js. Do not edit.\n`;

async function build({ check = false } = {}) {
  const src = await fs.readFile(SRC_PATH, "utf8");
  const output = GENERATED_HEADER + src;

  if (check) {
    let dist;
    try {
      dist = await fs.readFile(DIST_PATH, "utf8");
    } catch (error) {
      const message = error && error.code === "ENOENT"
        ? `Missing ${path.relative(process.cwd(), DIST_PATH)}. Run: node ${path.relative(process.cwd(), path.join(__dirname, "build.js"))}`
        : `Failed to read ${path.relative(process.cwd(), DIST_PATH)}: ${error?.message ?? error}`;
      throw new Error(message);
    }

    if (dist !== output) {
      throw new Error(
        `extensions/sample-hello/dist/extension.js is out of date. Run: node ${path.relative(
          process.cwd(),
          path.join(__dirname, "build.js")
        )}`
      );
    }
    return;
  }

  await fs.mkdir(path.dirname(DIST_PATH), { recursive: true });
  await fs.writeFile(DIST_PATH, output, "utf8");
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
  DIST_PATH,
  GENERATED_HEADER
};

