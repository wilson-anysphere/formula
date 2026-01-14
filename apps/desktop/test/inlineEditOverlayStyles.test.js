import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("InlineEditOverlay avoids display toggles via inline styles", () => {
  const overlayPath = path.join(__dirname, "..", "src", "ai", "inline-edit", "inlineEditOverlay.ts");
  const content = stripComments(fs.readFileSync(overlayPath, "utf8"));

  assert.equal(
    content.includes("style.display"),
    false,
    "InlineEditOverlay should not toggle visibility via `style.display`; use `hidden` / CSS classes instead",
  );
});
