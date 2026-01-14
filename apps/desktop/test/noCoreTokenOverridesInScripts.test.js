import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const srcRoot = path.join(desktopRoot, "src");

/**
 * @param {string} dirPath
 * @returns {string[]}
 */
function walkScriptFiles(dirPath) {
  /** @type {string[]} */
  const files = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      files.push(...walkScriptFiles(fullPath));
      continue;
    }
    if (!entry.isFile()) continue;
    if (!/\.[jt]sx?$/.test(entry.name)) continue;
    files.push(fullPath);
  }
  return files;
}

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop UI scripts should not override core design tokens (--space-*, --radius*, --motion-*, --font-*)", () => {
  const files = walkScriptFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    if (/\.(test|spec|vitest)\.[jt]sx?$/.test(rel)) return false;
    return true;
  });

  /** @type {string[]} */
  const violations = [];

  // Only match direct `.style.setProperty("--token", ...)` / `.style.removeProperty("--token")`
  // calls so we don't trigger on other string usages like `getPropertyValue("--space-4")`.
  const tokenOp =
    /\.style\.(?<op>setProperty|removeProperty)\(\s*(["'`])\s*(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)))\s*\2/gi;

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);
    const rel = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    tokenOp.lastIndex = 0;
    let match;
    while ((match = tokenOp.exec(stripped))) {
      const prop = match.groups?.prop ?? "";
      const op = match.groups?.op ?? "setProperty";
      const matchStart = match.index ?? 0;
      const matchText = match[0] ?? "";
      const propOffset = prop ? matchText.indexOf(prop) : -1;
      const propIndex = propOffset >= 0 ? matchStart + propOffset : matchStart;
      const line = getLineNumber(stripped, propIndex);
      violations.push(`${rel}:L${line}: ${op}(${prop}, ...)`);
    }
    tokenOp.lastIndex = 0;
  }

  assert.deepEqual(
    violations,
    [],
    `Found core token overrides in desktop UI scripts. Core design tokens must only be defined in src/styles/tokens.css:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
