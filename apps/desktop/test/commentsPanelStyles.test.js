import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

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

test("SpreadsheetApp comments panel/tooltip use CSS classes (no inline style.*)", () => {
  const filePath = path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts");
  const content = stripComments(fs.readFileSync(filePath, "utf8"));

  const togglePanel = extractBlock(content, "toggleCommentsPanel(): void");
  const createPanel = extractBlock(content, "private createCommentsPanel(");
  const createTooltip = extractBlock(content, "private createCommentTooltip(");
  const hideTooltip = extractBlock(content, "private hideCommentTooltip(): void");

  assert.equal(
    /\.style\./.test(togglePanel),
    false,
    "toggleCommentsPanel should not set inline styles; use a CSS class toggle (e.g. .comments-panel--visible)",
  );
  assert.equal(
    /\.style\./.test(createPanel),
    false,
    "createCommentsPanel should not set inline styles; move comment panel styling into src/styles/comments.css",
  );
  assert.equal(
    /\.style\./.test(createTooltip),
    false,
    "createCommentTooltip should not set inline styles; move comment tooltip styling into src/styles/comments.css",
  );
  assert.equal(
    /\.style\./.test(hideTooltip),
    false,
    "hideCommentTooltip should not set inline styles; use a CSS class toggle (e.g. .comment-tooltip--visible)",
  );

  // Keep show/hide behavior class-based (not style.display).
  assert.equal(
    content.includes("commentsPanel.style.display"),
    false,
    "Comments panel visibility should be controlled via CSS classes (not commentsPanel.style.display)",
  );
  assert.equal(
    content.includes("commentTooltip.style.display"),
    false,
    "Comment tooltip visibility should be controlled via CSS classes (not commentTooltip.style.display)",
  );

  // Panel/tooltip should be styled via semantic classes.
  for (const className of ["comments-panel", "comment-tooltip"]) {
    assert.ok(
      createPanel.includes(className) || createTooltip.includes(className),
      `Expected SpreadsheetApp to apply the ${className} CSS class`,
    );
  }

  // E2E tests and other automation rely on these data-testid hooks.
  for (const testId of [
    "comments-panel",
    "comments-active-cell",
    "new-comment-input",
    "submit-comment",
    "comment-tooltip",
  ]) {
    assert.ok(
      new RegExp(`dataset\\.testid\\s*=\\s*["']${testId}["']`).test(createPanel + createTooltip),
      `Expected createCommentsPanel/createCommentTooltip to preserve data-testid=\"${testId}\"`,
    );
  }

  // Sanity-check: class-based visibility selectors exist in CSS.
  const cssPath = path.join(__dirname, "..", "src", "styles", "comments.css");
  const css = fs.readFileSync(cssPath, "utf8");
  for (const selector of [".comments-panel--visible", ".comment-tooltip--visible"]) {
    assert.ok(css.includes(selector), `Expected comments.css to define ${selector}`);
  }
});
