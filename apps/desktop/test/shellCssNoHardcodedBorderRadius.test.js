import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const shellCssPath = path.join(desktopRoot, "src", "styles", "shell.css");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("shell.css should not hardcode border-radius pixel values (except 0 and 999px)", () => {
  const css = fs.readFileSync(shellCssPath, "utf8");
  // Avoid false positives in comments while keeping line numbers stable for error messages.
  const stripped = css.replace(/\/\*[\s\S]*?\*\//g, (comment) => comment.replace(/[^\n]/g, " "));

  /** @type {string[]} */
  const violations = [];
  const regex = /\bborder-radius\s*:\s*(\d+)px\b/gi;
  let match;
  while ((match = regex.exec(stripped))) {
    const value = Number(match[1]);
    if (value === 0 || value === 999) continue;
    violations.push(`L${getLineNumber(stripped, match.index)}: border-radius: ${match[1]}px`);
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius pixel values in shell.css. Use radius tokens (var(--radius*)), except for pills (999px) or 0:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
