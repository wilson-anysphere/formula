import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("tokens.css forced-colors overrides apply even when data-theme is set", () => {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = fs.readFileSync(tokensPath, "utf8");

  const mediaIndex = css.indexOf("@media (forced-colors: active), (prefers-contrast: more)");
  assert.ok(mediaIndex >= 0, "Expected tokens.css to define a forced-colors/prefers-contrast media query block");

  const afterMedia = css.slice(mediaIndex);

  // We can't fully evaluate CSS cascade in node tests, but we *can* ensure that
  // the forced-colors overrides match the specificity of `:root[data-theme=\"dark\"]`
  // so they win even when a concrete theme is set.
  assert.match(afterMedia, /:root,\s*\n\s*:root\[data-theme\]\s*\{/);

  // Guardrails: forced-colors should also neutralize common overlay tokens that
  // otherwise inherit from light/dark themes (shadows, tooltip colors, etc).
  assert.match(afterMedia, /--accent-hover\s*:/);
  assert.match(afterMedia, /--accent-active\s*:/);
  assert.match(afterMedia, /--panel-shadow\s*:/);
  assert.match(afterMedia, /--tooltip-bg\s*:/);
  assert.match(afterMedia, /--tooltip-text\s*:/);
  assert.match(afterMedia, /--dialog-shadow\s*:/);
  assert.match(afterMedia, /--dialog-backdrop\s*:/);
});
