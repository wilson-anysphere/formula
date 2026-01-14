import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("main.ts passes formatPainter exactly once to registerDesktopCommands", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const start = main.indexOf("registerDesktopCommands({");
  assert.notEqual(start, -1, "Expected main.ts to call registerDesktopCommands({ ... })");

  // `pageLayoutHandlers` comes immediately after `formatPainter` in the config object.
  // Guard against accidental duplicate `formatPainter:` keys, which TypeScript rejects (TS1117).
  const end = main.indexOf("pageLayoutHandlers", start);
  assert.notEqual(end, -1, "Expected main.ts to pass pageLayoutHandlers into registerDesktopCommands");

  const segment = main.slice(start, end);
  const matches = segment.match(/\bformatPainter\s*:/g) ?? [];
  assert.equal(
    matches.length,
    1,
    `Expected exactly one 'formatPainter:' property before pageLayoutHandlers (found ${matches.length})`,
  );
});

