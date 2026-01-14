import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

import { skipStringLiteral, stripComments } from "../../__tests__/sourceTextUtils";

function findMatchingDelimiter(source: string, openIndex: number, openChar: string, closeChar: string): number | null {
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    const ch = source[i];
    if (ch === "'" || ch === '"' || ch === "`") {
      i = skipStringLiteral(source, i) - 1;
      continue;
    }
    if (ch === openChar) depth += 1;
    else if (ch === closeChar) {
      depth -= 1;
      if (depth === 0) return i;
    }
  }
  return null;
}

function splitTopLevelArgs(argText: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let parens = 0;
  let braces = 0;
  let brackets = 0;

  for (let i = 0; i < argText.length; i += 1) {
    const ch = argText[i];
    if (ch === "'" || ch === '"' || ch === "`") {
      i = skipStringLiteral(argText, i) - 1;
      continue;
    }

    if (ch === "(") parens += 1;
    else if (ch === ")") parens = Math.max(0, parens - 1);
    else if (ch === "{") braces += 1;
    else if (ch === "}") braces = Math.max(0, braces - 1);
    else if (ch === "[") brackets += 1;
    else if (ch === "]") brackets = Math.max(0, brackets - 1);

    if (ch === "," && parens === 0 && braces === 0 && brackets === 0) {
      parts.push(argText.slice(start, i).trim());
      start = i + 1;
    }
  }

  const tail = argText.slice(start).trim();
  if (tail) parts.push(tail);
  return parts;
}

describe("oauthRedirectIpc wiring", () => {
  it("installs the oauth-redirect IPC readiness handshake in main.ts (prevents cold-start drops)", () => {
    // `main.ts` has many side effects and isn't safe to import in unit tests. Instead, validate
    // (lightly) that it wires the OAuth redirect listener and emits `oauth-redirect-ready`
    // *after* the listener is registered. The Rust host queues `oauth-redirect` URLs until this
    // handshake occurs; breaking it can drop deep-link redirects on cold start.
    const source = readFileSync(new URL("../../main.ts", import.meta.url), "utf8").replace(/\r\n?/g, "\n");
    // Strip comments from a bounded region around the oauth-redirect wiring rather than the entire
    // entrypoint. This keeps the guardrail resilient to unrelated patterns elsewhere in `main.ts`
    // (e.g. regex literals) while still preventing commented-out wiring from satisfying the test.
    const listenIndex = source.search(/\blisten\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["']/);
    expect(listenIndex).toBeGreaterThanOrEqual(0);
    const snippet = source.slice(Math.max(0, listenIndex - 2_000), Math.min(source.length, listenIndex + 6_000));
    const code = stripComments(snippet);

    // 1) Ensure we listen for Rust -> JS oauth redirect events.
    expect(code).toMatch(/\blisten\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["']/);

    // 2) Ensure we only emit readiness after the listener promise resolves.
    // Accept `const x = listen("oauth-redirect", ...); x.then(() => emit("oauth-redirect-ready"))`
    // and close equivalents (e.g. `await x; emit(...)`).
    // Count in the full entrypoint so we catch accidental early/unconditional emits elsewhere.
    const emitReadyMatches = Array.from(
      stripComments(source).matchAll(/\bemit\s*(?:\?\.)?\s*\(\s*["']oauth-redirect-ready["']/g),
    );
    expect(emitReadyMatches).toHaveLength(1);

    const escapeForRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const assignment =
      /^(?:\s*)(?:const|let|var)\s+(\w+)(?:\s*:\s*[^=]+)?\s*=\s*listen\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["']/m.exec(
        code,
      );
    if (assignment) {
      const varName = escapeForRegExp(assignment[1]!);
      const thenCall = new RegExp(String.raw`\b${varName}\b\s*(?:\?\.\s*then|\.\s*then)\b`, "m").exec(code);
      let hasThen = false;
      if (thenCall) {
        const searchStart = thenCall.index + thenCall[0].length;
        const openParen = code.indexOf("(", searchStart);
        if (openParen >= 0 && openParen - searchStart < 100) {
          const closeParen = findMatchingDelimiter(code, openParen, "(", ")");
          if (closeParen != null) {
            const argsText = code.slice(openParen + 1, closeParen);
            const [onFulfilled] = splitTopLevelArgs(argsText);
            hasThen =
              typeof onFulfilled === "string" &&
              (onFulfilled.includes("=>") || onFulfilled.includes("function")) &&
              /\bemit\s*(?:\?\.)?\s*\(\s*["']oauth-redirect-ready["']/.test(onFulfilled);
          }
        }
      }
      const hasAwait =
        new RegExp(
          String.raw`await\s+\b${varName}\b[\s\S]{0,750}?\bemit\s*(?:\?\.)?\s*\(\s*["']oauth-redirect-ready["']`,
          "m",
        ).test(code);
      expect(hasThen || hasAwait).toBe(true);
    } else {
      // Fallback for inlined promise chaining / top-level await.
      const hasThen =
        /\blisten\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["'][\s\S]{0,750}?\)\s*\.?\s*then\s*\(\s*(?:async\s*)?(?:\(\s*[^)]*\)\s*=>|\w+\s*=>|function\s*\(\s*[^)]*\))[\s\S]{0,750}?\bemit\s*(?:\?\.)?\s*\(\s*["']oauth-redirect-ready["']/.test(
          code,
        );
      const hasAwait =
        /await\s+listen\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["'][\s\S]{0,750}?\)\s*;?[\s\S]{0,750}?\bemit\s*(?:\?\.)?\s*\(\s*["']oauth-redirect-ready["']/.test(
          code,
        );
      expect(hasThen || hasAwait).toBe(true);
    }

    // 3) Ensure the listener forwards the payload into the OAuth broker.
    const listenCall = /\blisten\s*(?:<[^\n]*>)?\s*\(\s*["']oauth-redirect["']/.exec(code);
    expect(listenCall).toBeTruthy();
    if (listenCall) {
      const openParen = code.indexOf("(", listenCall.index);
      const closeParen = openParen >= 0 ? findMatchingDelimiter(code, openParen, "(", ")") : null;
      expect(closeParen).not.toBeNull();
      if (openParen >= 0 && closeParen != null) {
        const argsText = code.slice(openParen + 1, closeParen);
        const args = splitTopLevelArgs(argsText);
        expect(args.length).toBeGreaterThanOrEqual(2);
        expect(args[0]).toMatch(/["']oauth-redirect["']/);
        expect(args[1]).toMatch(/\boauthBroker\.observeRedirect\s*\(/);
      }
    }
  });
});
