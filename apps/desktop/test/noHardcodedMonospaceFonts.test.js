import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const srcRoot = path.join(desktopRoot, "src");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop styles should not hardcode monospace font stacks (use --font-mono token)", () => {
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

  const cssFiles = walkCssFiles(srcRoot)
    .filter((file) => {
      const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
      if (rel === "styles/tokens.css") return false;
      // Demo/sandbox assets are not part of the shipped UI bundle.
      if (rel.startsWith("grid/presence-renderer/")) return false;
      if (rel.includes("/demo/")) return false;
      if (rel.includes("/__tests__/")) return false;
      return true;
    })
    .sort((a, b) => a.localeCompare(b));

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;

  // Keep in sync with `--font-mono` in `src/styles/tokens.css`; we only allow hardcoded
  // monospace stacks in that single source of truth.
  //
  // Note: We intentionally scan only `font` / `font-family` declarations so we can preserve
  // quoted font family names like `"SF Mono"` without triggering false positives in unrelated
  // strings (e.g. `content: "monospace"` or urls).
  const forbiddenFontStackToken = /\b(?:ui-monospace|SFMono(?:-Regular)?|SF\s*Mono|Menlo|Consolas|monospace)\b/gi;

  /** @type {string[]} */
  const violations = [];

  for (const file of cssFiles) {
    const raw = fs.readFileSync(file, "utf8");
    const stripped = stripCssComments(raw);
    const relPath = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    cssDeclaration.lastIndex = 0;
    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = (decl?.groups?.prop ?? "").toLowerCase();
      if (prop !== "font-family" && prop !== "font") continue;

      const value = decl?.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      forbiddenFontStackToken.lastIndex = 0;
      let match;
      while ((match = forbiddenFontStackToken.exec(value))) {
        const absIndex = valueStart + (match.index ?? 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(`${relPath}:L${line}: ${match[0]}`);
      }
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
