import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");

test("Desktop CSS should not use brightness() filters (use tokens instead)", () => {
  const srcRoot = path.join(desktopRoot, "src");

  /**
   * @param {string} dirPath
   * @returns {string[]}
   */
  function walkCssFiles(dirPath) {
    /** @type {string[]} */
    const files = [];
    for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
      const fullPath = path.join(dirPath, entry.name);
      if (entry.isDirectory()) {
        files.push(...walkCssFiles(fullPath));
        continue;
      }
      if (!entry.isFile()) continue;
      if (!entry.name.endsWith(".css")) continue;
      files.push(fullPath);
    }
    return files;
  }

  const targets = walkCssFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    return true;
  });

  for (const target of targets) {
    const css = fs.readFileSync(target, "utf8");
    const stripped = stripCssNonSemanticText(css);
    assert.ok(
      !/\bbrightness\s*\(/i.test(stripped),
      `Expected ${path.relative(desktopRoot, target)} to avoid brightness(...) (use tokens instead)`,
    );
  }
});
