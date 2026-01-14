import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
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

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop UI should not hardcode border-radius pixel values (use radius tokens)", () => {
  const files = walkCssFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    return true;
  });
  /** @type {string[]} */
  const violations = [];

  for (const file of files) {
    const css = fs.readFileSync(file, "utf8");
    // Avoid false positives in comments while keeping line numbers stable for error messages.
    const stripped = css.replace(/\/\*[\s\S]*?\*\//g, (comment) => comment.replace(/[^\n]/g, " "));

    const declRegex = /\bborder-radius\s*:\s*([^;}]*)/gi;
    let declMatch;
    while ((declMatch = declRegex.exec(stripped))) {
      const value = declMatch[1] ?? "";
      // `declMatch[0]` ends with the captured group, so this points at the first character of the value.
      const valueStart = declMatch.index + declMatch[0].length - value.length;

      const pxRegex = /(\d+(?:\.\d+)?)px\b/g;
      let pxMatch;
      while ((pxMatch = pxRegex.exec(value))) {
        const px = Number(pxMatch[1]);
        if (px === 0) continue;

        const absIndex = valueStart + pxMatch.index;
        const line = getLineNumber(stripped, absIndex);
        violations.push(
          `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: border-radius: ${pxMatch[1]}px`,
        );
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius pixel values in desktop UI styles. Use radius tokens (var(--radius*)), except for 0:\n${violations
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
