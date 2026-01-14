import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

function skipStringLiteral(source: string, start: number): number {
  const quote = source[start];
  if (quote !== "'" && quote !== '"' && quote !== "`") return start + 1;
  let i = start + 1;
  while (i < source.length) {
    const ch = source[i];
    if (ch === "\\") {
      i += 2;
      continue;
    }
    if (ch === quote) return i + 1;
    i += 1;
  }
  return source.length;
}

function stripComments(source: string): string {
  // Remove JS comments without accidentally stripping `https://...` inside string literals.
  // This is intentionally lightweight: it's not a full parser, but is sufficient for guardrail
  // matching in `main.ts` and avoids treating commented-out wiring as valid.
  let out = "";
  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];

    if (ch === "'" || ch === '"' || ch === "`") {
      const end = skipStringLiteral(source, i);
      out += source.slice(i, end);
      i = end - 1;
      continue;
    }

    if (ch === "/" && source[i + 1] === "/") {
      // Line comment.
      i += 2;
      while (i < source.length && source[i] !== "\n") i += 1;
      if (i < source.length) out += "\n";
      continue;
    }

    if (ch === "/" && source[i + 1] === "*") {
      // Block comment (preserve newlines so we don't accidentally join tokens across lines).
      i += 2;
      while (i < source.length) {
        const next = source[i];
        if (next === "\n") out += "\n";
        if (next === "*" && source[i + 1] === "/") {
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }

    out += ch;
  }
  return out;
}

describe("openFileIpc wiring", () => {
  it("installs the open-file IPC handshake in main.ts (prevents cold-start drops)", () => {
    // `main.ts` has a lot of side effects and isn't safe to import in unit tests. Instead,
    // validate (lightly) that it wires the open-file IPC helper responsible for emitting
    // `open-file-ready` *after* the listener is registered.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");
    const code = stripComments(source);

    // Ensure the helper is imported (as an actual import statement, not just mentioned in a comment).
    // Allow optional `.ts`/`.js` extensions so this guardrail doesn't break on harmless specifier refactors.
    expect(code).toMatch(
      /^\s*import\s+\{[^}]*\binstallOpenFileIpc\b[^}]*\}\s+from\s+["']\.\/tauri\/openFileIpc(?:\.(?:ts|js))?["']/m,
    );

    // Ensure the helper is actually used. This guards against a regression where the helper
    // remains in the tree but the startup wiring is removed.
    expect(code).toMatch(/^\s*(?:void\s+)?installOpenFileIpc\(\s*\{[\s\S]*?\blisten\b[\s\S]*?\bemit\b[\s\S]*?\bonOpenPath\b/m);
  });
});
