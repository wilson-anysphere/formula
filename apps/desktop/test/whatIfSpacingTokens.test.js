import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

test("what-if styles keep spacing on the shared --space-* scale", () => {
  const cssPath = path.join(desktopRoot, "src", "styles", "what-if.css");
  const css = stripCssNonSemanticText(fs.readFileSync(cssPath, "utf8"));

  const declRe = /(?:^|[;{])\s*(gap|padding(?:-[a-z]+)?|margin(?:-[a-z]+)?)\s*:\s*([^;}]+)/g;

  /** @type {{ prop: string; value: string }[]} */
  const offenders = [];

  let match;
  while ((match = declRe.exec(css))) {
    const prop = match[1];
    const value = match[2].trim();
    if (/\b\d+(?:\.\d+)?px\b/.test(value)) {
      offenders.push({ prop, value });
    }
  }

  assert.equal(
    offenders.length,
    0,
    `Found pixel-based spacing declarations in what-if.css:\n${offenders.map((o) => `- ${o.prop}: ${o.value}`).join("\n")}`,
  );
});
