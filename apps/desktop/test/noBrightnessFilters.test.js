import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");

test("Accent hover styles should not use filter: brightness(...)", () => {
  const targets = [
    path.join(desktopRoot, "src", "styles", "ribbon.css"),
    path.join(desktopRoot, "src", "styles", "ui.css"),
    path.join(desktopRoot, "src", "styles", "sort-filter.css"),
    path.join(desktopRoot, "src", "titlebar", "titlebar.css"),
  ];

  for (const target of targets) {
    const css = fs.readFileSync(target, "utf8");
    assert.ok(
      !/filter\s*:\s*brightness\(/i.test(css),
      `Expected ${path.relative(desktopRoot, target)} to avoid filter: brightness(...) (use tokens instead)`,
    );
  }
});
