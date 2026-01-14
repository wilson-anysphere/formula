import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, it } from "vitest";

import { CLIPBOARD_LIMITS } from "./provider.js";
import { stripRustComments } from "../../__tests__/sourceTextUtils";

function readRustConstExpression(source: string, constName: string): string {
  const pattern = new RegExp(`pub const\\s+${constName}\\s*:\\s*usize\\s*=\\s*([^;]+);`);
  const match = source.match(pattern);
  if (!match) {
    throw new Error(
      [
        `Unable to locate Rust clipboard constant ${constName}.`,
        "This test enforces that JS clipboard caps stay in sync with the Rust backend.",
        "If the Rust file layout changed, update the parser in:",
        "  apps/desktop/src/clipboard/platform/provider.limits.vitest.ts",
      ].join("\n")
    );
  }
  return match[1]!.trim();
}

function evalRustUsizeExpression(expr: string): number {
  const sanitized = expr.replaceAll("_", "").trim();
  if (!/^[0-9\s+*/()%\-]+$/.test(sanitized)) {
    throw new Error(
      [
        `Unsupported Rust const expression: "${expr}".`,
        "This sync test currently supports simple numeric arithmetic expressions (e.g. `5 * 1024 * 1024`).",
        "Either rewrite the Rust const into a numeric expression, or update the evaluator in:",
        "  apps/desktop/src/clipboard/platform/provider.limits.vitest.ts",
      ].join("\n")
    );
  }

  // Safe: `sanitized` contains only digits + arithmetic punctuation.
  // eslint-disable-next-line no-new-func
  const value = Function(`"use strict"; return (${sanitized});`)() as unknown;

  if (typeof value !== "number" || !Number.isFinite(value) || !Number.isSafeInteger(value)) {
    throw new Error(`Failed to evaluate Rust const expression "${expr}" into a safe integer.`);
  }

  return value;
}

describe("clipboard caps stay in sync with the Rust backend", () => {
  it("JS CLIPBOARD_LIMITS matches Rust MAX_* constants", () => {
    const rustPath = fileURLToPath(new URL("../../../src-tauri/src/clipboard/mod.rs", import.meta.url));
    const rustSource = stripRustComments(readFileSync(rustPath, "utf8"));

    const rustMaxPngBytes = evalRustUsizeExpression(readRustConstExpression(rustSource, "MAX_PNG_BYTES"));
    const rustMaxTextBytes = evalRustUsizeExpression(readRustConstExpression(rustSource, "MAX_TEXT_BYTES"));
    const rustMaxPlainTextWriteBytes = evalRustUsizeExpression(
      readRustConstExpression(rustSource, "MAX_PLAINTEXT_WRITE_BYTES"),
    );

    const mismatches: string[] = [];

    if (CLIPBOARD_LIMITS.maxImageBytes !== rustMaxPngBytes) {
      mismatches.push(
        `- JS provider CLIPBOARD_LIMITS.maxImageBytes (${CLIPBOARD_LIMITS.maxImageBytes}) does not match Rust MAX_PNG_BYTES (${rustMaxPngBytes}).`
      );
    }

    if (CLIPBOARD_LIMITS.maxRichTextBytes !== rustMaxTextBytes) {
      mismatches.push(
        `- JS provider CLIPBOARD_LIMITS.maxRichTextBytes (${CLIPBOARD_LIMITS.maxRichTextBytes}) does not match Rust MAX_TEXT_BYTES (${rustMaxTextBytes}).`
      );
    }

    if (CLIPBOARD_LIMITS.maxPlainTextWriteBytes !== rustMaxPlainTextWriteBytes) {
      mismatches.push(
        `- JS provider CLIPBOARD_LIMITS.maxPlainTextWriteBytes (${CLIPBOARD_LIMITS.maxPlainTextWriteBytes}) does not match Rust MAX_PLAINTEXT_WRITE_BYTES (${rustMaxPlainTextWriteBytes}).`
      );
    }

    if (mismatches.length > 0) {
      throw new Error(
        [
          "Clipboard byte limits drift detected (keep these files in sync):",
          "  - apps/desktop/src/clipboard/platform/provider.js",
          "  - apps/desktop/src-tauri/src/clipboard/mod.rs",
          "",
          ...mismatches,
        ].join("\n")
      );
    }
  });
});
