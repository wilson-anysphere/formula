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

test("desktop UI scripts should not use brightness() filters (use tokens instead)", () => {
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
  const brightnessRe = /\bbrightness\s*\(/i;

  /** @type {{ re: RegExp, kind: string }[]} */
  const patterns = [
    // CSS style strings (e.g. `style: "filter: brightness(0.9);"`).
    { re: /\bfilter\s*:\s*(?<value>[^;"'`]*)/gi, kind: "filter" },
    // React style objects (e.g. `{ filter: "brightness(0.9)" }`).
    { re: /\bfilter\s*:\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "filter (style object)" },
    // DOM style assignment (e.g. `el.style.filter = "brightness(0.9)"`)
    { re: /\.style\.filter\s*=\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.filter" },
    // setProperty("filter", "brightness(0.9)")
    {
      re: /\.style\.setProperty\(\s*(["'])filter\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
      kind: "setProperty(filter)",
    },
    // setAttribute("style", "filter: brightness(0.9)")
    {
      re: /\bsetAttribute\(\s*(["'])style\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
      kind: "setAttribute(style)",
    },
    // cssText assignment (e.g. `el.style.cssText = "filter: brightness(0.9)"`)
    { re: /\.style\.cssText\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.cssText" },
  ];

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);
    const rel = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    for (const { re, kind } of patterns) {
      re.lastIndex = 0;
      let match;
      while ((match = re.exec(stripped))) {
        const value = match.groups?.value ?? "";
        if (!brightnessRe.test(value)) continue;

        const matchStart = match.index ?? 0;
        const matchText = match[0] ?? "";
        const valueOffset = matchText.indexOf(value);
        const absIndex = matchStart + (valueOffset >= 0 ? valueOffset : 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(`${rel}:L${line}: ${kind} = ${JSON.stringify(value)}`);
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found brightness() filters in desktop UI scripts. Use tokens instead:\n${violations.map((v) => `- ${v}`).join("\n")}`,
  );
});
