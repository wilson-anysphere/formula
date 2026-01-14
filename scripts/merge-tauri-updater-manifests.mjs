#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";

import { mergeTauriUpdaterManifests, normalizeVersion } from "./tauri-updater-manifest.mjs";

export { mergeTauriUpdaterManifests, normalizeVersion };

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function usage() {
  return [
    "Merge multiple Tauri updater `latest.json` manifests into one multi-platform manifest.",
    "",
    "Usage:",
    "  node scripts/merge-tauri-updater-manifests.mjs --out <output.json> <input1.json> <input2.json> [...]",
    "",
    "Notes:",
    "  - Versions must match after normalization (`vX.Y.Z` vs `X.Y.Z`).",
    "  - Platform keys are merged by union; conflicts fail.",
    "",
  ].join("\n");
}

/**
 * @param {string[]} argv
 */
export async function main(argv = process.argv.slice(2)) {
  if (argv.includes("--help") || argv.includes("-h")) {
    console.log(usage());
    return 0;
  }

  const outIdx = argv.indexOf("--out");
  if (outIdx < 0 || outIdx + 1 >= argv.length) {
    console.error("Missing --out <path> argument.\n");
    console.error(usage());
    return 1;
  }

  const outPath = argv[outIdx + 1];
  const inputPaths = argv
    .slice(0, outIdx)
    .concat(argv.slice(outIdx + 2))
    .filter((arg) => arg && !arg.startsWith("-"));

  if (inputPaths.length === 0) {
    console.error("Expected at least one input manifest path.\n");
    console.error(usage());
    return 1;
  }

  const manifests = [];
  for (const relOrAbs of inputPaths) {
    const full = path.isAbsolute(relOrAbs) ? relOrAbs : path.join(repoRoot, relOrAbs);
    const raw = await readFile(full, "utf8");
    manifests.push(JSON.parse(raw));
  }

  const merged = mergeTauriUpdaterManifests(manifests);
  const fullOut = path.isAbsolute(outPath) ? outPath : path.join(repoRoot, outPath);
  await writeFile(fullOut, `${JSON.stringify(merged, null, 2)}\n`, "utf8");

  console.log(
    `Merged ${manifests.length} manifest(s) into ${path.relative(repoRoot, fullOut)} (version ${merged.version}, ${Object.keys(merged.platforms).length} platform(s)).`,
  );
  return 0;
}

const isMain = (() => {
  const argv1 = process.argv[1];
  if (!argv1) return false;
  try {
    const invoked = pathToFileURL(path.resolve(argv1)).href;
    return invoked === import.meta.url;
  } catch {
    return false;
  }
})();

if (isMain) {
  main().then((code) => {
    process.exitCode = code;
  });
}

