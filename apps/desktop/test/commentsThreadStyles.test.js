import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractBlock(source, startNeedle) {
  const start = source.indexOf(startNeedle);
  assert.ok(start !== -1, `Expected to find ${startNeedle}`);

  const firstBrace = source.indexOf("{", start);
  assert.ok(firstBrace !== -1, `Expected ${startNeedle} to include an opening {`);

  let depth = 0;
  for (let i = firstBrace; i < source.length; i++) {
    const ch = source[i];
    if (ch === "{") depth += 1;
    if (ch === "}") depth -= 1;
    if (depth === 0) {
      return source.slice(start, i + 1);
    }
  }

  assert.fail(`Failed to find matching closing brace for ${startNeedle}`);
}

test("SpreadsheetApp.renderCommentThread is styled via CSS classes (no inline style.*)", () => {
  const filePath = path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts");
  const content = fs.readFileSync(filePath, "utf8");
  const fn = extractBlock(content, "private renderCommentThread(");

  assert.equal(
    /\.style\./.test(fn),
    false,
    "renderCommentThread should not set inline styles; move comment thread UI styling into src/styles/comments.css",
  );

  // Sanity-check that the thread is still tagged with a CSS class for styling.
  assert.ok(
    /comment-thread/.test(fn),
    "renderCommentThread should apply a .comment-thread* class so styling lives in comments.css",
  );

  // E2E tests rely on these data-testid hooks.
  for (const testId of ["comment-thread", "resolve-comment", "reply-input", "submit-reply"]) {
    assert.ok(
      new RegExp(`dataset\\.testid\\s*=\\s*["']${testId}["']`).test(fn),
      `renderCommentThread should preserve data-testid=\"${testId}\" for Playwright`,
    );
  }

  // Resolved visual state should be driven by CSS (data-resolved selector), not JS inline styles.
  const cssPath = path.join(__dirname, "..", "src", "styles", "comments.css");
  const css = fs.readFileSync(cssPath, "utf8");
  assert.ok(
    css.includes('.comment-thread[data-resolved="true"]'),
    "comments.css should style resolved threads via .comment-thread[data-resolved=\"true\"]",
  );
});
