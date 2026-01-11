import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

const extensionPackage = require("../shared/extension-package/index.js");
const signing = require("../shared/crypto/signing.js");

function usage() {
  // eslint-disable-next-line no-console
  console.error(
    [
      "Usage:",
      "  pnpm extension:pack <dir> --out <file> [--private-key <pemPath>]",
      "  pnpm extension:verify <file> --pubkey <pemPath>",
      "  pnpm extension:inspect <file>",
      "",
    ].join("\n"),
  );
}

function readFlag(args, name) {
  const idx = args.indexOf(name);
  if (idx === -1) return null;
  return args[idx + 1] || null;
}

function readFlagOrEnv(args, name, envKey) {
  const direct = readFlag(args, name);
  if (direct) return direct;
  const fromEnv = envKey ? process.env[envKey] : null;
  return fromEnv || null;
}

async function readPemFromArg(value) {
  if (!value) return null;
  if (value.includes(path.sep) || value.endsWith(".pem")) {
    return fs.readFile(value, "utf8");
  }
  return value;
}

async function cmdPack(args) {
  const dir = args[0] ? path.resolve(args[0]) : null;
  const out = readFlagOrEnv(args, "--out", "npm_config_out");
  const privateKeyArg =
    readFlagOrEnv(args, "--private-key", "npm_config_private_key") ?? process.env.FORMULA_EXTENSION_PRIVATE_KEY;

  if (!dir || !out) {
    usage();
    process.exit(1);
  }

  let privateKeyPem = await readPemFromArg(privateKeyArg);
  let publicKeyPem = null;

  if (!privateKeyPem) {
    const generated = signing.generateEd25519KeyPair();
    privateKeyPem = generated.privateKeyPem;
    publicKeyPem = generated.publicKeyPem;
    // eslint-disable-next-line no-console
    console.error("No --private-key provided; generated an ephemeral Ed25519 keypair for signing.");
    // eslint-disable-next-line no-console
    console.error("Public key (save this to verify):\n" + publicKeyPem.trim());
  }

  const bytes = await extensionPackage.createExtensionPackage(dir, { formatVersion: 2, privateKeyPem });
  await fs.mkdir(path.dirname(out), { recursive: true });
  await fs.writeFile(out, bytes);
}

async function cmdVerify(args) {
  const file = args[0] ? path.resolve(args[0]) : null;
  const pubkeyArg = readFlagOrEnv(args, "--pubkey", "npm_config_pubkey") ?? process.env.FORMULA_EXTENSION_PUBLIC_KEY;
  if (!file || !pubkeyArg) {
    usage();
    process.exit(1);
  }

  const publicKeyPem = await readPemFromArg(pubkeyArg);
  const bytes = await fs.readFile(file);
  const formatVersion = extensionPackage.detectExtensionPackageFormatVersion(bytes);

  if (formatVersion === 1) {
    throw new Error("v1 packages do not embed signatures. Verify using the detached signature from the marketplace.");
  }

  const verified = extensionPackage.verifyExtensionPackageV2(bytes, publicKeyPem);
  // eslint-disable-next-line no-console
  console.log(
    JSON.stringify(
      {
        ok: true,
        formatVersion,
        id: `${verified.manifest.publisher}.${verified.manifest.name}`,
        version: verified.manifest.version,
        fileCount: verified.fileCount,
        unpackedSize: verified.unpackedSize,
      },
      null,
      2,
    ),
  );
}

async function cmdInspect(args) {
  const file = args[0] ? path.resolve(args[0]) : null;
  if (!file) {
    usage();
    process.exit(1);
  }

  const bytes = await fs.readFile(file);
  const formatVersion = extensionPackage.detectExtensionPackageFormatVersion(bytes);

  if (formatVersion === 1) {
    const pkg = extensionPackage.readExtensionPackageV1(bytes);
    // eslint-disable-next-line no-console
    console.log(
      JSON.stringify(
        {
          formatVersion,
          manifest: pkg.manifest,
          fileCount: pkg.files.length,
          paths: pkg.files.map((f) => f.path),
        },
        null,
        2,
      ),
    );
    return;
  }

  const pkg = extensionPackage.readExtensionPackageV2(bytes);
  // eslint-disable-next-line no-console
  console.log(
    JSON.stringify(
      {
        formatVersion,
        manifest: pkg.manifest,
        checksums: pkg.checksums,
        signature: pkg.signature,
        fileCount: pkg.files.size,
        paths: [...pkg.files.keys()],
      },
      null,
      2,
    ),
  );
}

async function main() {
  const argv = process.argv.slice(2);
  const cmd = argv[0];
  const args = argv.slice(1);

  try {
    switch (cmd) {
      case "pack":
        await cmdPack(args);
        return;
      case "verify":
        await cmdVerify(args);
        return;
      case "inspect":
        await cmdInspect(args);
        return;
      default:
        usage();
        process.exit(1);
    }
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error(error?.message ?? String(error));
    process.exit(1);
  }
}

// Ensure stack traces point at this file when invoked via pnpm.
void fileURLToPath(import.meta.url);
main();
