import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { stripHtmlComments } from "../apps/desktop/test/sourceTextUtils.js";

const DOC_PATH = path.resolve(process.cwd(), "docs/05-ai-integration.md");

/**
 * Extract a markdown section by heading text.
 *
 * @param {string} markdown
 * @param {string} headingLine exact line including leading `### `
 */
function extractSection(markdown, headingLine) {
  const lines = markdown.split("\n");
  const startLine = lines.findIndex((line) => line === headingLine);
  assert.notEqual(startLine, -1, `Expected docs to contain heading: ${headingLine}`);

  /** @type {string[]} */
  const out = [];

  // Keep this intentionally simple: stop at the next heading at the same or higher
  // level (### / ## / #), but ignore headings inside fenced code blocks.
  let inFence = false;
  for (let i = startLine; i < lines.length; i += 1) {
    const line = lines[i];
    if (i > startLine && !inFence && /^#{1,3}\s/.test(line)) break;

    out.push(line);

    if (line.startsWith("```")) inFence = !inFence;
  }

  return out.join("\n");
}

test("docs: RAG Over Cells describes deterministic HashEmbedder embeddings (not user-configurable)", async () => {
  // Strip HTML comments so commented-out docs cannot satisfy or fail assertions.
  const markdown = stripHtmlComments(await fs.readFile(DOC_PATH, "utf8"));
  const section = extractSection(markdown, "### RAG Over Cells");
  const lower = section.toLowerCase();

  // Positive assertions: keep the docs explicit about the current implementation.
  assert.match(lower, /not user-configurable/);
  assert.match(section, /HashEmbedder/);
  assert.match(lower, /deterministic hash embeddings/);
  assert.match(lower, /does not accept user api keys/);
  assert.match(lower, /local model[\s>]+setup/);
  assert.match(lower, /cursor[- ]managed embedding service/);

  // Negative assertions: prevent regressions where provider-specific embedding
  // setup instructions creep back into the RAG docs.
  for (const forbidden of ["open" + "ai", "an" + "thropic", "ol" + "lama"]) {
    assert.equal(lower.includes(forbidden), false, `RAG docs must not mention '${forbidden}'`);
  }
});
