import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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
    "error",
    "warning",
    "success"
  ];

  for (const token of required) {
    assert.ok(defined.has(token), `Expected tokens.css to define --${token}`);
  }
});
