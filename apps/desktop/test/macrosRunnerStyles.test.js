import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("renderMacroRunner is class-driven (no inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "macros", "dom_ui.ts");
  const src = fs.readFileSync(srcPath, "utf8");

  assert.equal(
    /\.style\./.test(src) || /\.style\[/.test(src),
    false,
    "renderMacroRunner should not set inline styles (element.style.*); use src/styles/macros-runner.css classes instead",
  );

  const requiredClasses = [
    "macros-runner",
    "macros-runner__header",
    "macros-runner__security-banner",
    "macros-runner__controls",
    "macros-runner__select",
    "macros-runner__button",
    "macros-runner__output",
  ];

  for (const className of requiredClasses) {
    assert.ok(src.includes(className), `Expected dom_ui.ts to reference CSS class "${className}"`);
  }
});

