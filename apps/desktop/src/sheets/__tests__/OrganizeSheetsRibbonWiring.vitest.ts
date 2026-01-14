import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

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

describe("Organize Sheets ribbon wiring", () => {
  it("routes the ribbon command id to openOrganizeSheets()", () => {
    const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
    const source = stripComments(readFileSync(mainTsPath, "utf8"));

    // Ensure the ribbon command id is explicitly handled and opens the dialog.
    // Be tolerant of minor formatting differences (single vs double quotes, whitespace).
    const caseMatch = source.match(/case\s+["']home\.cells\.format\.organizeSheets["']\s*:/);
    expect(caseMatch).not.toBeNull();
    const caseIndex = caseMatch?.index ?? -1;
    expect(caseIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(caseIndex, caseIndex + 300)).toMatch(/openOrganizeSheets\s*\(/);

    // Ensure the helper exists and delegates to `openOrganizeSheetsDialog`.
    const fnMatch = source.match(/(?:function\s+openOrganizeSheets\s*\(|const\s+openOrganizeSheets\s*=\s*\(\)\s*=>)/);
    expect(fnMatch).not.toBeNull();
    const fnIndex = fnMatch?.index ?? -1;
    expect(fnIndex).toBeGreaterThanOrEqual(0);
    expect(source.slice(fnIndex, fnIndex + 1600)).toContain("openOrganizeSheetsDialog(");
  });
});
