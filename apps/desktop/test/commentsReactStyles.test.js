import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("React comments components avoid inline styles (except tooltip positioning)", () => {
  const commentsPanelPath = path.join(__dirname, "..", "src", "comments", "CommentsPanel.tsx");
  const tooltipPath = path.join(__dirname, "..", "src", "comments", "CommentTooltip.tsx");

  const panel = stripComments(fs.readFileSync(commentsPanelPath, "utf8"));
  const tooltip = stripComments(fs.readFileSync(tooltipPath, "utf8"));

  assert.equal(
    /\bstyle\s*=/.test(panel),
    false,
    "CommentsPanel.tsx should not use inline styles; styling should live in src/styles/comments.css",
  );

  // Sanity-check that the panel is still wired with the shared CSS classes.
  for (const className of ["comments-panel-view", "comment-thread"]) {
    assert.ok(
      panel.includes(className),
      `CommentsPanel.tsx should include the ${className} class to share comments.css styling`,
    );
  }

  // The tooltip needs dynamic positioning, but should still use the shared CSS class.
  assert.ok(
    tooltip.includes("comment-tooltip"),
    "CommentTooltip.tsx should include the .comment-tooltip class so styling lives in comments.css",
  );
});
