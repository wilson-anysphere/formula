#!/usr/bin/env node

const path = require("node:path");

const { buildExtension, checkExtension } = require("./builder");

function usage() {
  // eslint-disable-next-line no-console
  console.error(
    [
      "Usage:",
      "  formula-extension-builder build <extensionDir> [--watch] [--minify] [--sourcemap] [--strict] [--write-manifest] [--entry <path>]",
      "  formula-extension-builder check <extensionDir> [--minify] [--sourcemap] [--strict] [--entry <path>]",
      "",
    ].join("\n"),
  );
}

function hasFlag(args, name) {
  return args.includes(name);
}

function readFlag(args, name) {
  const idx = args.indexOf(name);
  if (idx === -1) return null;
  return args[idx + 1] || null;
}

async function main() {
  const args = process.argv.slice(2);
  const cmd = args[0];
  const extensionDir = args[1] ? path.resolve(args[1]) : null;

  if (!cmd || !extensionDir || (cmd !== "build" && cmd !== "check")) {
    usage();
    process.exit(1);
  }

  const options = {
    watch: hasFlag(args, "--watch"),
    minify: hasFlag(args, "--minify"),
    sourcemap: hasFlag(args, "--sourcemap"),
    strict: hasFlag(args, "--strict"),
    writeManifest: hasFlag(args, "--write-manifest"),
    entry: readFlag(args, "--entry"),
  };

  if (cmd === "build") {
    await buildExtension(extensionDir, options);
    return;
  }

  await checkExtension(extensionDir, options);
}

main().catch((error) => {
  // eslint-disable-next-line no-console
  console.error(error?.message ?? String(error));
  process.exit(1);
});

