#!/usr/bin/env node
/**
 * Excel function parity report (code-driven).
 *
 * This script intentionally avoids extra dependencies so it can run in CI or in a
 * fresh checkout with only Node installed.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "../..");

function readUtf8(relPath) {
  return fs.readFileSync(path.join(repoRoot, relPath), "utf8");
}

function readJson(relPath) {
  return JSON.parse(readUtf8(relPath));
}

function parseFtabNames(ftabSource) {
  // Find the initializer for: `pub const FTAB: ... = [ ... ];`
  const decl = "pub const FTAB:";
  const declIdx = ftabSource.indexOf(decl);
  if (declIdx === -1) {
    throw new Error(`Could not find \`${decl}\` in crates/formula-biff/src/ftab.rs`);
  }

  const eqIdx = ftabSource.indexOf("=", declIdx);
  if (eqIdx === -1) {
    throw new Error("Could not find '=' after `pub const FTAB:`");
  }

  const openIdx = ftabSource.indexOf("[", eqIdx);
  if (openIdx === -1) {
    throw new Error("Could not find '[' starting FTAB array initializer");
  }

  const closeIdx = ftabSource.indexOf("];", openIdx);
  if (closeIdx === -1) {
    throw new Error("Could not find closing '];' for FTAB array initializer");
  }

  const body = ftabSource.slice(openIdx + 1, closeIdx);
  const names = [];
  const stringLiteralRe = /"((?:\\.|[^"\\])*)"/g;
  for (const match of body.matchAll(stringLiteralRe)) {
    // We do not fully unescape Rust string literals here. FTAB names are expected to be plain ASCII.
    names.push(match[1]);
  }
  return names;
}

function usage() {
  // Keep usage text simple so it stays readable in CI logs.
  return [
    "Usage:",
    "  node tools/parity/report_functions.mjs [--list-missing]",
    "",
    "Options:",
    "  --list-missing   Print FTAB function names that are missing from shared/functionCatalog.json",
  ].join("\n");
}

const args = new Set(process.argv.slice(2));
if (args.has("--help") || args.has("-h")) {
  console.log(usage());
  process.exit(0);
}

const catalog = readJson("shared/functionCatalog.json");
if (!catalog || !Array.isArray(catalog.functions)) {
  throw new Error("shared/functionCatalog.json: expected top-level { functions: [...] }");
}

const implementedNames = catalog.functions.map((f) => String(f.name).toUpperCase());
const implementedSet = new Set(implementedNames);

const ftabSource = readUtf8("crates/formula-biff/src/ftab.rs");
const ftabNamesAll = parseFtabNames(ftabSource);
const ftabNames = ftabNamesAll.filter((name) => name.length > 0).map((n) => n.toUpperCase());
const ftabSet = new Set(ftabNames);

const missingFromCatalog = [...ftabSet].filter((name) => !implementedSet.has(name)).sort();

console.log("Excel function parity (code-driven)");
console.log(`Implemented functions (shared/functionCatalog.json): ${implementedSet.size}`);
console.log(`BIFF FTAB function names (non-empty): ${ftabSet.size}`);
console.log(`FTAB names missing from engine catalog (approx): ${missingFromCatalog.length}`);

if (args.has("--list-missing")) {
  console.log("");
  for (const name of missingFromCatalog) {
    console.log(name);
  }
}

