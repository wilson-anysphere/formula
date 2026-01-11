import { promises as fs } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { KeyRing } from "../../../packages/security/crypto/keyring.js";

export type KeyRingJson = {
  currentVersion: number;
  keys: Record<string, string>;
};

export function generateKeyRingJson(): KeyRingJson {
  return KeyRing.create().toJSON();
}

export function rotateKeyRingJson(input: unknown): KeyRingJson {
  const ring = KeyRing.fromJSON(input);
  ring.rotate();
  return ring.toJSON();
}

export function validateKeyRingJson(input: unknown): {
  currentVersion: number;
  availableVersions: number[];
} {
  const ring = KeyRing.fromJSON(input);
  const availableVersions = [...ring.keysByVersion.keys()].sort((a, b) => a - b);
  return { currentVersion: ring.currentVersion, availableVersions };
}

function usage(): string {
  return [
    "Usage:",
    "  keyring generate",
    "  keyring rotate --in <path> [--out <path>]",
    "  keyring validate --in <path>",
    "",
    "Notes:",
    "  - rotate keeps existing key versions so historical data remains decryptable.",
    "  - Output is KeyRing JSON ({ currentVersion, keys }).",
  ].join("\n");
}

function takeFlag(args: string[], flag: string): string | undefined {
  const idx = args.indexOf(flag);
  if (idx === -1) return undefined;
  const value = args[idx + 1];
  if (!value || value.startsWith("-")) {
    throw new Error(`Missing value for ${flag}`);
  }
  return value;
}

async function atomicWriteFile(filePath: string, contents: string): Promise<void> {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath);
  const tmpPath = path.join(dir, `.${base}.${process.pid}.${Date.now()}.tmp`);

  await fs.writeFile(tmpPath, contents, "utf8");
  try {
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "EEXIST" || code === "EPERM") {
      await fs.rm(filePath, { force: true });
      await fs.rename(tmpPath, filePath);
      return;
    }
    throw err;
  }
}

async function readJsonFile(filePath: string): Promise<unknown> {
  const raw = await fs.readFile(filePath, "utf8");
  return JSON.parse(raw);
}

async function main(argv: string[]): Promise<void> {
  const args = [...argv];
  if (args[0] === "keyring") args.shift();

  const cmd = args.shift();
  if (!cmd || cmd === "--help" || cmd === "-h") {
    process.stdout.write(`${usage()}\n`);
    return;
  }

  if (args.includes("--help") || args.includes("-h")) {
    process.stdout.write(`${usage()}\n`);
    return;
  }

  if (cmd === "generate") {
    process.stdout.write(`${JSON.stringify(generateKeyRingJson(), null, 2)}\n`);
    return;
  }

  if (cmd === "rotate") {
    const inPath = takeFlag(args, "--in");
    const outPath = takeFlag(args, "--out");
    if (!inPath) {
      throw new Error("rotate requires --in <path>");
    }

    const input = await readJsonFile(inPath);
    const rotated = rotateKeyRingJson(input);
    const output = `${JSON.stringify(rotated, null, 2)}\n`;

    if (outPath) {
      await atomicWriteFile(outPath, output);
    } else {
      process.stdout.write(output);
    }
    return;
  }

  if (cmd === "validate") {
    const inPath = takeFlag(args, "--in");
    if (!inPath) {
      throw new Error("validate requires --in <path>");
    }

    const input = await readJsonFile(inPath);
    const summary = validateKeyRingJson(input);
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
    return;
  }

  throw new Error(`Unknown command: ${cmd}\n\n${usage()}`);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? "").href) {
  try {
    await main(process.argv.slice(2));
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
}
