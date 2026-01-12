import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const DOC_PATH = path.resolve(process.cwd(), "docs/05-ai-integration.md");

/**
 * Extract a markdown section by heading text.
 *
 * @param {string} markdown
 * @param {string} headingLine exact line including leading `### `
 */
function extractSection(markdown, headingLine) {
  const start = markdown.indexOf(headingLine);
  assert.notEqual(start, -1, `Expected docs to contain heading: ${headingLine}`);

  // Find the next top-level section heading at the same depth (`### `).
  const afterStart = start + headingLine.length;
  const next = markdown.indexOf("\n### ", afterStart);
  const end = next === -1 ? markdown.length : next;
  return markdown.slice(start, end);
}

test("docs: RAG Over Cells describes deterministic HashEmbedder embeddings (not user-configurable)", async () => {
  const markdown = await fs.readFile(DOC_PATH, "utf8");
  const section = extractSection(markdown, "### RAG Over Cells");
  const lower = section.toLowerCase();

  // Positive assertions: keep the docs explicit about the current implementation.
  assert.match(lower, /not user-configurable/);
  assert.match(section, /HashEmbedder/);
  assert.match(lower, /deterministic hash embeddings/);
  assert.match(lower, /does not accept user api keys/);
  assert.match(lower, /local model[\s>]+setup/);

  // Negative assertions: prevent regressions where provider-specific embedding
  // setup instructions creep back into the RAG docs.
  for (const forbidden of ["open" + "ai", "an" + "thropic", "ol" + "lama"]) {
    assert.equal(lower.includes(forbidden), false, `RAG docs must not mention '${forbidden}'`);
  }
});
