import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const stylesRoot = path.join(desktopRoot, "src", "styles");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop styles should not hardcode monospace font stacks (use --font-mono token)", () => {
  const cssFiles = fs
    .readdirSync(stylesRoot, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".css") && entry.name !== "tokens.css")
    .map((entry) => path.join(stylesRoot, entry.name))
    .sort((a, b) => a.localeCompare(b));

  // Keep in sync with `--font-mono` in `src/styles/tokens.css`; we only allow hardcoded
  // monospace stacks in that single source of truth.
  const forbiddenFontStackToken = /\b(ui-monospace|SFMono|Menlo|Consolas|monospace)\b/gi;

  /** @type {string[]} */
  const violations = [];

  for (const file of cssFiles) {
    const raw = fs.readFileSync(file, "utf8");
    const stripped = stripCssNonSemanticText(raw);

    let match;
    while ((match = forbiddenFontStackToken.exec(stripped))) {
      const line = getLineNumber(stripped, match.index ?? 0);
      violations.push(
        `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${match[0]}`,
      );
    }

    forbiddenFontStackToken.lastIndex = 0;
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded monospace font stacks in desktop styles. Use var(--font-mono) from src/styles/tokens.css:\n${violations
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});

