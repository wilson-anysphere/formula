import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

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

test("tokens.css defines required design tokens", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = fs.readFileSync(tokensPath, "utf8");

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

test("tokens.css uses tight radius tokens (4px/3px) per mockups", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = fs.readFileSync(tokensPath, "utf8");

  const rootMatch = css.match(/:root\s*\{([\s\S]*?)\}/);
  assert.ok(rootMatch, "tokens.css missing :root block");

  const vars = parseVarsFromBlock(rootMatch[1]);
  assert.equal(vars["radius"], "4px");
  assert.equal(vars["radius-sm"], "3px");
});
