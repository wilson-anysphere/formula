import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const stylesRoot = path.join(desktopRoot, "src", "styles");

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

test("core desktop styles should not hardcode border-radius: 2px (use --radius-xs)", () => {
  const files = walkCssFiles(stylesRoot);
  /** @type {string[]} */
  const violations = [];

  for (const file of files) {
    const css = fs.readFileSync(file, "utf8");
    // Avoid false positives in comments.
    const stripped = css.replace(/\/\*[\s\S]*?\*\//g, " ");
    if (/\bborder-radius\s*:\s*2px\b/i.test(stripped)) {
      violations.push(path.relative(desktopRoot, file));
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius: 2px in desktop core styles. Use var(--radius-xs) instead:\n${violations
      .map((file) => `- ${file}`)
      .join("\n")}`,
  );
});

