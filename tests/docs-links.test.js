import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { stripHtmlComments } from "../apps/desktop/test/sourceTextUtils.js";

async function listFilesRecursively(dir) {
  /** @type {string[]} */
  const out = [];
  const entries = await fs.readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await listFilesRecursively(full)));
    } else {
      out.push(full);
    }
  }
  return out;
}

function stripFencedCodeBlocks(markdown) {
  const lines = markdown.split("\n");
  /** @type {string[]} */
  const out = [];

  let inFence = false;
  let fenceMarker = null;

  for (const line of lines) {
    const trimmed = line.trimStart();
    const isFence = trimmed.startsWith("```") || trimmed.startsWith("~~~");
    if (isFence) {
      const marker = trimmed.startsWith("```") ? "```" : "~~~";
      if (!inFence) {
        inFence = true;
        fenceMarker = marker;
      } else if (marker === fenceMarker) {
        inFence = false;
        fenceMarker = null;
      }
      continue;
    }

    if (!inFence) out.push(line);
  }

  return out.join("\n");
}

test("docs: markdown links resolve to existing files", async () => {
  const docsDir = path.resolve(process.cwd(), "docs");
  const files = (await listFilesRecursively(docsDir)).filter((p) => p.endsWith(".md"));

  /** @type {{ file: string; link: string; resolved: string }[]} */
  const broken = [];

  // Simple markdown link matcher. We intentionally keep this conservative:
  // - we ignore links inside fenced code blocks
  // - we ignore http(s) and mailto links
  // - we verify local paths (with optional #fragment / ?query stripped)
  const linkRe = /\[[^\]]*\]\(([^)]+)\)/g;

  for (const file of files) {
    const raw = await fs.readFile(file, "utf8");
    // Strip HTML comments so commented-out markdown links cannot satisfy or fail assertions.
    const markdown = stripFencedCodeBlocks(stripHtmlComments(raw));

    for (const match of markdown.matchAll(linkRe)) {
      let url = match[1]?.trim() ?? "";

      // Markdown allows `<...>` around destinations (often used for paths with spaces).
      if (url.startsWith("<") && url.endsWith(">")) {
        url = url.slice(1, -1).trim();
      }

      if (
        url.startsWith("http://") ||
        url.startsWith("https://") ||
        url.startsWith("mailto:")
      ) {
        continue;
      }
      if (url.startsWith("#")) continue;

      const withoutFragment = url.split("#", 1)[0].split("?", 1)[0];
      if (!withoutFragment) continue;

      const resolved = withoutFragment.startsWith("/")
        ? path.resolve(process.cwd(), withoutFragment.slice(1))
        : path.resolve(path.dirname(file), withoutFragment);

      try {
        await fs.stat(resolved);
      } catch {
        broken.push({ file, link: url, resolved });
      }
    }
  }

  if (broken.length > 0) {
    const msg =
      "Broken markdown links:\n" +
      broken
        .map(
          ({ file, link, resolved }) =>
            `- ${path.relative(process.cwd(), file)} -> ${link} (expected at ${path.relative(
              process.cwd(),
              resolved
            )})`
        )
        .join("\n");
    assert.fail(msg);
  }
});
