import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssComments } from "./sourceTextUtils.js";
import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function parseVarsFromBlock(blockBody) {
  /** @type {Record<string, string>} */
  const vars = {};
  const regex = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let match = null;
  while ((match = regex.exec(blockBody))) {
    vars[match[1]] = match[2].trim();
  }
  return vars;
}

/**
 * @param {string} css
 * @param {string} name
 */
function collectVarAssignments(css, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const regex = new RegExp(`--${escaped}\\s*:\\s*([^;]+);`, "gi");
  /** @type {string[]} */
  const values = [];
  let match = null;
  while ((match = regex.exec(css))) {
    values.push(match[1].trim());
  }
  return values;
}

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

test("tokens.css defines required design tokens", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = stripCssComments(fs.readFileSync(tokensPath, "utf8"));

  const defined = new Set();
  const regex = /--([a-z0-9-]+)\s*:/gi;
  let match = null;
  while ((match = regex.exec(css))) {
    defined.add(match[1]);
  }

  const required = [
    "font-sans",
    "font-mono",
    "space-0",
    "space-1",
    "space-2",
    "space-3",
    "space-4",
    "space-5",
    "space-6",
    "radius",
    "radius-sm",
    "radius-xs",
    "radius-pill",
    "bg-primary",
    "bg-secondary",
    "bg-tertiary",
    "text-primary",
    "text-secondary",
    "border",
    "accent",
    "accent-hover",
    "accent-active",
    "accent-light",
    "link",
    "error",
    "warning",
    "success"
  ];

  for (const token of required) {
    assert.ok(defined.has(token), `Expected tokens.css to define --${token}`);
  }
});

test("tokens.css defines the shared spacing scale (2px base) per mockups", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = fs.readFileSync(tokensPath, "utf8");

  const rootMatch = css.match(/:root\s*\{([\s\S]*?)\}/);
  assert.ok(rootMatch, "tokens.css missing :root block");

  const vars = parseVarsFromBlock(rootMatch[1]);
  assert.equal(vars["space-0"], "0px");
  assert.equal(vars["space-1"], "2px");
  assert.equal(vars["space-2"], "4px");
  assert.equal(vars["space-3"], "6px");
  assert.equal(vars["space-4"], "8px");
  assert.equal(vars["space-5"], "12px");
  assert.equal(vars["space-6"], "16px");
});

test("tokens.css uses tight radius tokens (4px/3px) per mockups", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = stripCssComments(fs.readFileSync(tokensPath, "utf8"));

  const rootMatch = css.match(/:root\s*\{([\s\S]*?)\}/);
  assert.ok(rootMatch, "tokens.css missing :root block");

  const vars = parseVarsFromBlock(rootMatch[1]);
  assert.equal(vars["radius"], "4px");
  assert.equal(vars["radius-sm"], "3px");
  assert.equal(vars["radius-xs"], "2px");
  assert.equal(vars["radius-pill"], "999px");
});

test("space tokens stay consistent across themes (no accidental overrides)", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = stripCssComments(fs.readFileSync(tokensPath, "utf8"));

  /** @type {Record<string, string>} */
  const expected = {
    "space-0": "0px",
    "space-1": "2px",
    "space-2": "4px",
    "space-3": "6px",
    "space-4": "8px",
    "space-5": "12px",
    "space-6": "16px",
  };

  for (const [token, value] of Object.entries(expected)) {
    const values = collectVarAssignments(css, token);
    assert.ok(values.length > 0, `Expected tokens.css to define --${token}`);
    for (const actual of values) {
      assert.equal(actual, value, `Expected --${token} to always be ${value} (got ${actual})`);
    }
  }
});

test("core design tokens (--space-*, --radius*, --font-*) are only defined in tokens.css", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const files = walkCssFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    if (rel === "styles/tokens.css") return false;
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    return true;
  });

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>--[-\w]+)\s*:/gi;
  /** @type {string[]} */
  const violations = [];

  for (const file of files) {
    const css = fs.readFileSync(file, "utf8");
    const stripped = stripCssNonSemanticText(css);
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");

    cssDeclaration.lastIndex = 0;
    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (
        !prop.startsWith("--space-") &&
        !prop.startsWith("--radius") &&
        prop !== "--font-sans" &&
        prop !== "--font-mono"
      ) {
        continue;
      }
      const line = getLineNumber(stripped, decl.index ?? 0);
      violations.push(`${rel}:L${line}: ${prop}`);
    }
    cssDeclaration.lastIndex = 0;
  }

  assert.deepEqual(
    violations,
    [],
    `Found core token overrides outside src/styles/tokens.css:\n${violations
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});

test("radius tokens stay consistent across themes (no accidental overrides)", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = stripCssComments(fs.readFileSync(tokensPath, "utf8"));

  /** @type {Record<string, string>} */
  const expected = {
    radius: "4px",
    "radius-sm": "3px",
    "radius-xs": "2px",
    "radius-pill": "999px",
  };

  for (const [token, value] of Object.entries(expected)) {
    const values = collectVarAssignments(css, token);
    assert.ok(values.length > 0, `Expected tokens.css to define --${token}`);
    for (const actual of values) {
      assert.equal(actual, value, `Expected --${token} to always be ${value} (got ${actual})`);
    }
  }
});
