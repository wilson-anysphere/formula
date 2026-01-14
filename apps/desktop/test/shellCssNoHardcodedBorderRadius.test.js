import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const shellCssPath = path.join(desktopRoot, "src", "styles", "shell.css");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("shell.css should not hardcode border-radius values (except 0)", () => {
  const css = fs.readFileSync(shellCssPath, "utf8");
  const stripped = stripCssNonSemanticText(css);

  /** @type {string[]} */
  const violations = [];

  const declRegex = /\bborder(?:-(?:top|bottom|start|end)-(?:left|right|start|end))?-radius\s*:\s*([^;}]*)/gi;
  let declMatch;
  while ((declMatch = declRegex.exec(stripped))) {
    const value = declMatch[1] ?? "";
    const valueStart = declMatch.index + declMatch[0].length - value.length;

    const unitRegex =
      /([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|%|rem|em|vh|vw|vmin|vmax|cm|mm|in|pt|pc|ch|ex)(?![A-Za-z0-9_])/gi;
    let unitMatch;
    while ((unitMatch = unitRegex.exec(value))) {
      const numeric = unitMatch[1];
      const unit = unitMatch[2] ?? "";
      const n = Number(numeric);
      if (!Number.isFinite(n)) continue;
      if (n === 0) continue;

      const absIndex = valueStart + unitMatch.index;
      violations.push(`L${getLineNumber(stripped, absIndex)}: border-radius: ${numeric}${unit}`);
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius values in shell.css. Use radius tokens (var(--radius*)), except for 0:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
