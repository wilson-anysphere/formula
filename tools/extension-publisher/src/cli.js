#!/usr/bin/env node

const path = require("node:path");

const { publishExtension } = require("./publisher");

function usage() {
  // eslint-disable-next-line no-console
  console.error(
    [
      "Usage:",
      "  formula-extension-publisher publish <extensionDir> --marketplace <url> --token <token> --private-key <pemPath>",
      "",
    ].join("\n"),
  );
}

async function main() {
  const args = process.argv.slice(2);
  const cmd = args[0];
  if (cmd !== "publish") {
    usage();
    process.exit(1);
  }

  const extensionDir = args[1] ? path.resolve(args[1]) : null;

  function readFlag(name) {
    const idx = args.indexOf(name);
    if (idx === -1) return null;
    return args[idx + 1] || null;
  }

  const marketplaceUrl = readFlag("--marketplace");
  const token = readFlag("--token");
  const privateKey = readFlag("--private-key");

  if (!extensionDir || !marketplaceUrl || !token || !privateKey) {
    usage();
    process.exit(1);
  }

  const published = await publishExtension({
    extensionDir,
    marketplaceUrl,
    token,
    privateKeyPemOrPath: privateKey,
  });

  // eslint-disable-next-line no-console
  console.log(`Published ${published.id}@${published.version}`);
}

main().catch((error) => {
  // eslint-disable-next-line no-console
  console.error(error);
  process.exit(1);
});

