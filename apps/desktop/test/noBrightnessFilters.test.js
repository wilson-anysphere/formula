import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");

test("Desktop CSS should not use brightness() filters (use tokens instead)", () => {
  const stylesDir = path.join(desktopRoot, "src", "styles");
  const styleTargets = fs
    .readdirSync(stylesDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".css"))
    .map((entry) => path.join(stylesDir, entry.name));

  const targets = [
    ...styleTargets,
    // Accent-driven titlebar hover styling is still token-based but lives outside src/styles.
    path.join(desktopRoot, "src", "titlebar", "titlebar.css"),
  ];

  const uniqueTargets = [...new Set(targets)];

  for (const target of uniqueTargets) {
    const css = fs.readFileSync(target, "utf8");
    assert.ok(
      !/\bbrightness\s*\(/i.test(css),
      `Expected ${path.relative(desktopRoot, target)} to avoid brightness(...) (use tokens instead)`,
    );
  }
});
