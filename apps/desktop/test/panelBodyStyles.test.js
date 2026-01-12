import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractBlock(source, selector) {
  const start = source.indexOf(selector);
  assert.ok(start !== -1, `Expected to find selector ${selector}`);

  const firstBrace = source.indexOf("{", start);
  assert.ok(firstBrace !== -1, `Expected ${selector} to include an opening {`);

  let depth = 0;
  for (let i = firstBrace; i < source.length; i++) {
    const ch = source[i];
    if (ch === "{") depth += 1;
    if (ch === "}") depth -= 1;
    if (depth === 0) {
      return source.slice(start, i + 1);
    }
  }

  assert.fail(`Failed to find matching closing brace for ${selector}`);
}

test("workspace.css defines panel body container/fill helpers (class-driven mounts)", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "workspace.css");
  const css = fs.readFileSync(cssPath, "utf8");

  for (const legacy of [".dock-panel__mount", ".panel-mount--fill-column", ".dock-panel__body--fill"]) {
    assert.equal(
      css.includes(legacy),
      false,
      `workspace.css should not define legacy dock mount/fill helper ${legacy}; prefer panel-body__container/panel-body--fill`,
    );
  }

  const containerBlock = extractBlock(css, ".panel-body__container");
  assert.match(containerBlock, /\bmin-width\s*:\s*0\b/, "panel-body__container should set min-width: 0");
  assert.match(containerBlock, /\bmin-height\s*:\s*0\b/, "panel-body__container should set min-height: 0");
  assert.match(containerBlock, /\bflex\s*:\s*1\b/, "panel-body__container should set flex: 1");
  assert.match(containerBlock, /\bheight\s*:\s*100%/, "panel-body__container should set height: 100%");
  assert.match(containerBlock, /\bdisplay\s*:\s*flex\b/, "panel-body__container should set display: flex");
  assert.match(containerBlock, /\bflex-direction\s*:\s*column\b/, "panel-body__container should set flex-direction: column");

  const fillBlock = extractBlock(css, ".panel-body--fill");
  assert.match(fillBlock, /\bflex\s*:\s*1\b/, "panel-body--fill should set flex: 1");
  assert.match(fillBlock, /\bmin-height\s*:\s*0\b/, "panel-body--fill should set min-height: 0");
  assert.match(fillBlock, /\bpadding\s*:\s*0\b/, "panel-body--fill should set padding: 0");
  assert.match(fillBlock, /\bdisplay\s*:\s*flex\b/, "panel-body--fill should set display: flex");
  assert.match(fillBlock, /\bflex-direction\s*:\s*column\b/, "panel-body--fill should set flex-direction: column");
  assert.match(fillBlock, /\bcolor\s*:\s*var\(--text-primary\)/, "panel-body--fill should default text color to var(--text-primary)");
});
