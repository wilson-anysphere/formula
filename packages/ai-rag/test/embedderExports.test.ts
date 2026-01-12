import { expect, test } from "vitest";

import * as aiRag from "../src/index.js";

test("ai-rag does not export remote/local provider embedders", () => {
  const remoteName = "Open" + "AI" + "Embedder";
  const localName = "Ollama" + "Embedder";

  expect(Object.prototype.hasOwnProperty.call(aiRag, remoteName)).toBe(false);
  expect(Object.prototype.hasOwnProperty.call(aiRag, localName)).toBe(false);
});

