import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("desktop index.html is style-free (no inline <style> or style=\"\" attrs)", () => {
  const htmlPath = path.join(__dirname, "..", "index.html");
  const html = fs.readFileSync(htmlPath, "utf8");

  const styleTags = [...html.matchAll(/<style\b[^>]*>[\s\S]*?<\/style>/gi)];
  assert.equal(
    styleTags.length,
    0,
    "apps/desktop/index.html should not include inline <style> blocks; move shell styling into src/styles/*.css",
  );

  // Disallow inline styles on elements. If we need tiny, critical styling in the
  // future (e.g. to avoid a flash of unstyled content), prefer a small dedicated
  // CSS file instead so it stays token-driven and testable.
  const styleAttrs = [...html.matchAll(/\sstyle\s*=\s*["'][^"']*["']/gi)];
  assert.equal(
    styleAttrs.length,
    0,
    "apps/desktop/index.html should not include inline style=\"...\" attributes; move styling into src/styles/*.css",
  );
});

