import assert from "node:assert/strict";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

async function collectDtsFiles(dir, out = []) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectDtsFiles(full, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".d.ts")) continue;
    out.push(full);
  }
  return out;
}

function isIdentChar(ch) {
  // A simplified JS/TS identifier character check. This intentionally ignores unicode
  // identifier categories since our `.d.ts` sources are ASCII.
  return /[A-Za-z0-9_$]/.test(ch);
}

function containsAnyTypeToken(code) {
  /** @type {Array<{ kind: "code" | "template" | "templateExpr" | "string" | "lineComment" | "blockComment", quote?: string, depth?: number }>} */
  const stack = [{ kind: "code" }];

  for (let i = 0; i < code.length; ) {
    const top = stack[stack.length - 1];
    const ch = code[i];
    const next = code[i + 1];

    if (top.kind === "lineComment") {
      if (ch === "\n") stack.pop();
      i += 1;
      continue;
    }

    if (top.kind === "blockComment") {
      if (ch === "*" && next === "/") {
        stack.pop();
        i += 2;
      } else {
        i += 1;
      }
      continue;
    }

    if (top.kind === "string") {
      if (ch === "\\") {
        // Skip escaped char.
        i += 2;
        continue;
      }
      if (ch === top.quote) {
        stack.pop();
        i += 1;
        continue;
      }
      i += 1;
      continue;
    }

    if (top.kind === "template") {
      if (ch === "\\") {
        i += 2;
        continue;
      }
      if (ch === "`") {
        stack.pop();
        i += 1;
        continue;
      }
      // Template literal type expressions: scan inside `${ ... }`.
      if (ch === "$" && next === "{") {
        stack.push({ kind: "templateExpr", depth: 1 });
        i += 2;
        continue;
      }
      i += 1;
      continue;
    }

    // `code` or `templateExpr`
    if (ch === "/" && next === "/") {
      stack.push({ kind: "lineComment" });
      i += 2;
      continue;
    }
    if (ch === "/" && next === "*") {
      stack.push({ kind: "blockComment" });
      i += 2;
      continue;
    }
    if (ch === "'" || ch === '"') {
      stack.push({ kind: "string", quote: ch });
      i += 1;
      continue;
    }
    if (ch === "`") {
      stack.push({ kind: "template" });
      i += 1;
      continue;
    }

    if (top.kind === "templateExpr") {
      if (ch === "{") {
        top.depth += 1;
        i += 1;
        continue;
      }
      if (ch === "}") {
        top.depth -= 1;
        i += 1;
        if (top.depth === 0) stack.pop();
        continue;
      }
    }

    // Detect bare `any` tokens outside strings/comments.
    if (ch === "a" && code.slice(i, i + 3) === "any") {
      const prev = i > 0 ? code[i - 1] : "";
      const after = code[i + 3] ?? "";
      if (!isIdentChar(prev) && !isIdentChar(after)) return true;
    }

    i += 1;
  }
  return false;
}

test("ai-context .d.ts files do not use the `any` type", async () => {
  const srcDir = fileURLToPath(new URL("../src", import.meta.url));
  const files = await collectDtsFiles(srcDir);

  const offenders = [];
  for (const file of files) {
    const text = await readFile(file, "utf8");
    if (containsAnyTypeToken(text)) offenders.push(file.slice(srcDir.length + 1));
  }

  offenders.sort();
  assert.deepEqual(offenders, []);
});
