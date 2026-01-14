import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("SelectionRenderer fill handle uses --formula-grid-selection-handle token", () => {
  const filePath = path.join(__dirname, "..", "src", "selection", "renderer.ts");
  const source = fs.readFileSync(filePath, "utf8");

  assert.match(
    source,
    /fillHandleColor:\s*resolveToken\(\"--formula-grid-selection-handle\"/,
    "Expected SelectionRenderer theme defaults to resolve --formula-grid-selection-handle",
  );

  assert.match(
    source,
    /ctx\.fillStyle\s*=\s*style\.fillHandleColor/,
    "Expected SelectionRenderer to paint the fill handle using fillHandleColor",
  );
});

