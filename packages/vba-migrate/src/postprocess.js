import { a1ToRowCol, rowColToA1 } from "./a1.js";
/**
 * This module is used in both Node (CLI/tests) and browser environments
 * (desktop/webview). Keep Node builtins behind dynamic imports so bundlers don't
 * choke on `node:*` specifiers.
 */

function isNodeRuntime() {
  // Treat jsdom/webview environments as "browser" even though they execute inside
  // a Node process, because `child_process` is not available in the real target
  // runtime (Tauri webview).
  return typeof window === "undefined" && typeof process !== "undefined" && !!process.versions?.node;
}

let nodeDepsPromise = null;

async function loadNodeDeps() {
  if (nodeDepsPromise) return nodeDepsPromise;
  nodeDepsPromise = (async () => {
    const fs = await import(/* @vite-ignore */ "node:fs");
    const os = await import(/* @vite-ignore */ "node:os");
    const path = await import(/* @vite-ignore */ "node:path");
    const child = await import(/* @vite-ignore */ "node:child_process");
    const util = await import(/* @vite-ignore */ "node:util");
    const execFileAsync = util.promisify(child.execFile);
    return {
      writeFileSync: fs.writeFileSync,
      tmpdir: os.tmpdir,
      join: path.join,
      execFileAsync
    };
  })();
  return nodeDepsPromise;
}

function browserSafeErrorMessage(error) {
  if (!error) return "Unknown error";
  if (error instanceof Error && typeof error.message === "string") return error.message;
  // eslint-disable-next-line @typescript-eslint/no-base-to-string
  return String(error);
}

function stripMarkdownCodeFences(text) {
  const trimmed = String(text || "").trim();
  const fenceMatch = /^```[a-zA-Z0-9_-]*\n([\s\S]*?)\n```$/m.exec(trimmed);
  if (!fenceMatch) return trimmed;
  return fenceMatch[1].trim();
}

function ensurePythonWrapper(code) {
  const cleaned = String(code || "").trim();
  const hasMain = /\bdef\s+main\s*\(/.test(cleaned);
  if (hasMain) {
    const hasEntrypoint = /if\s+__name__\s*==\s*(['"])__main__\1\s*:/.test(cleaned);
    if (hasEntrypoint) return cleaned;
    return `${cleaned}\n\nif __name__ == "__main__":\n    main()`;
  }

  // Wrap "loose" script bodies in a main() to make execution consistent.
  const bodyLines = cleaned.split(/\r?\n/);
  const indented = bodyLines.map((line) => (line.trim() ? `    ${line}` : "")).join("\n");

  return `def main():\n${indented}\n\nif __name__ == "__main__":\n    main()`;
}

function ensurePythonImportFormula(code) {
  const cleaned = String(code || "").trim();
  if (/\bimport\s+formula\b/.test(cleaned)) return cleaned;
  return `import formula\n\n${cleaned}`;
}

function normalizePythonObjectModel(code) {
  let out = String(code || "");

  // Common LLM artifact: leaving VBA-ish property casing.
  out = out.replace(/\.Value\b/g, "");
  out = out.replace(/\.Formula\b/g, ".formula");

  // Common artifact: using Range() method as if in VBA.
  // e.g. sheet.Range("A1") -> sheet["A1"]
  out = out.replace(/\bsheet\.(?:Range|range)\(\s*(['"])([^'"]+)\1\s*\)/g, 'sheet["$2"]');

  // Common artifact: using Cells(row, col) method as if in VBA.
  // e.g. sheet.Cells(1,2) -> sheet["B1"]
  out = out.replace(/\bsheet\.Cells\(\s*(\d+)\s*,\s*(\d+)\s*\)/gi, (_match, row, col) => {
    try {
      const addr = rowColToA1(Number(row), Number(col));
      return `sheet["${addr}"]`;
    } catch {
      return _match;
    }
  });

  out = out.replace(/\bActiveSheet\b/g, "formula.active_sheet");
  return out;
}

function ensureTypeScriptWrapper(code) {
  const cleaned = String(code || "").trim();
  if (/\bexport\s+default\s+async\s+function\s+main\b/.test(cleaned)) return cleaned;

  const bodyLines = cleaned.split(/\r?\n/);
  const indented = bodyLines.map((line) => (line.trim() ? `  ${line}` : "")).join("\n");
  return `export default async function main(ctx) {\n${indented}\n}`;
}

function normalizeTypeScriptObjectModel(code) {
  let out = String(code || "");

  // Convert common VBA-ish method casing into the ScriptRuntime API (`getRange` / `getCell`).
  out = out.replace(/\b([A-Za-z_][A-Za-z0-9_]*)\.(?:Range|range)\(/g, "$1.getRange(");

  // VBA `Cells(row, col)` is 1-indexed. Prefer converting to A1 + getRange(...) so we don't
  // have to guess whether ScriptRuntime's `getCell(row, col)` is 0- or 1-indexed.
  out = out.replace(
    /\b([A-Za-z_][A-Za-z0-9_]*)\.(?:Cells|cells)\(\s*(\d+)\s*,\s*(\d+)\s*\)/g,
    (_match, ident, row, col) => {
      try {
        const addr = rowColToA1(Number(row), Number(col));
        return `${ident}.getRange("${addr}")`;
      } catch {
        return _match;
      }
    },
  );

  // Legacy helper used by early prompts/tests: `sheet.cell(1,2)` (also 1-indexed).
  out = out.replace(
    /\b([A-Za-z_][A-Za-z0-9_]*)\.cell\(\s*(\d+)\s*,\s*(\d+)\s*\)/gi,
    (_match, ident, row, col) => {
      try {
        const addr = rowColToA1(Number(row), Number(col));
        return `${ident}.getRange("${addr}")`;
      } catch {
        return _match;
      }
    },
  );

  // Convert property assignment patterns into async ScriptRuntime method calls.
  // Examples:
  //   sheet.getRange("A1").Value = 1;
  //   sheet.getRange("A1").value = 1;
  //   sheet.getRange("A3").Formula = "=A1+B1";
  // -> await sheet.getRange("A1").setValue(1);
  // -> await sheet.getRange("A3").setFormulas([["=A1+B1"]]);
  const lines = out.split(/\r?\n/);
  const rewritten = lines.map((line) => {
    const match = /^(\s*)(.+?)\.(value|Value|formula|Formula)\s*=\s*(.+?)\s*;?\s*$/.exec(line);
    if (!match) return line;
    const indent = match[1];
    const target = match[2];
    const prop = match[3].toLowerCase();
    const rhs = match[4];

    const rangeMatch = /\.getRange\(\s*(['"])(?<addr>[^'"]+)\1\s*\)/.exec(target);
    const addr = rangeMatch?.groups?.addr ?? null;
    const isMultiCell = typeof addr === "string" ? addr.includes(":") : false;

    const rangeDims = (address) => {
      if (typeof address !== "string") return { rows: 1, cols: 1 };
      const [startRaw, endRaw] = address.split(":");
      const start = a1ToRowCol(startRaw);
      const end = endRaw ? a1ToRowCol(endRaw) : start;
      const rows = Math.abs(end.row - start.row) + 1;
      const cols = Math.abs(end.col - start.col) + 1;
      return { rows, cols };
    };

    const isScalarLiteral = (expr) => {
      const trimmed = String(expr || "").trim();
      if (!trimmed) return false;
      if (trimmed.startsWith("[")) return false;
      if (trimmed.startsWith("{")) return false;
      if (trimmed.startsWith('"') || trimmed.startsWith("'")) return true;
      if (/^(true|false|null|undefined)$/i.test(trimmed)) return true;
      if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return true;
      return false;
    };

    const fillMatrixLiteral = (expr, { rows, cols }) => {
      const row = Array.from({ length: cols }, () => expr).join(", ");
      const matrix = Array.from({ length: rows }, () => `[${row}]`).join(", ");
      return `[${matrix}]`;
    };

    const fillMatrixExpr = (expr, { rows, cols }) => {
      if (String(expr || "").trim().startsWith("[")) return expr;
      if (isScalarLiteral(expr)) return fillMatrixLiteral(expr, { rows, cols });
      return `Array.from({ length: ${rows} }, () => Array(${cols}).fill(${expr}))`;
    };

    if (prop === "value") {
      const method = isMultiCell ? "setValues" : "setValue";
      const nextRhs = isMultiCell ? fillMatrixExpr(rhs, rangeDims(addr)) : rhs;
      return `${indent}await ${target}.${method}(${nextRhs});`;
    }

    if (prop === "formula") {
      if (isMultiCell) {
        const nextRhs = fillMatrixExpr(rhs, rangeDims(addr));
        return `${indent}await ${target}.setFormulas(${nextRhs});`;
      }
      const rhsTrimmed = rhs.trim();
      if (rhsTrimmed.startsWith("[")) {
        return `${indent}await ${target}.setFormulas(${rhs});`;
      }
      return `${indent}await ${target}.setFormulas([[${rhs}]]);`;
    }

    return line;
  });

  return rewritten.join("\n");
}

export async function postProcessGeneratedCode({ code, target }) {
  const stripped = stripMarkdownCodeFences(code);
  if (target === "python") {
    let python = stripped;
    python = normalizePythonObjectModel(python);
    python = ensurePythonWrapper(python);
    python = ensurePythonImportFormula(python);
    return python.trim() + "\n";
  }

  if (target === "typescript") {
    let ts = stripped;
    ts = normalizeTypeScriptObjectModel(ts);
    ts = ensureTypeScriptWrapper(ts);
    return ts.trim() + "\n";
  }

  throw new Error(`Unknown target: ${target}`);
}

export async function validateGeneratedCodeCompiles({ code, target }) {
  if (target === "python") {
    // Validate via `py_compile` so we catch indentation/syntax errors deterministically.
    if (isNodeRuntime()) {
      const { writeFileSync, tmpdir, join, execFileAsync } = await loadNodeDeps();
      const tmpDir = tmpdir();
      const filePath = join(tmpDir, `vba-migrate-${Date.now()}-${Math.random().toString(16).slice(2)}.py`);
      writeFileSync(filePath, code, "utf8");
      try {
        await execFileAsync("python", ["-m", "py_compile", filePath]);
        return { ok: true };
      } catch (error) {
        return { ok: false, error: error?.stderr?.toString?.() ?? browserSafeErrorMessage(error) };
      }
    }

    // Browser/webview environments do not ship with `python` available, so the
    // compile check is best-effort. We still run post-processing to normalize
    // common artifacts, but skip the external compiler step.
    return { ok: true, skipped: true };
  }

  if (target === "typescript") {
    // We intentionally restrict generated code to TS that is also valid JS/ESM.
    // In Node we use `node --check`; in browsers we fall back to a lightweight
    // parse of the generated script body.
    if (isNodeRuntime()) {
      const { writeFileSync, tmpdir, join, execFileAsync } = await loadNodeDeps();
      const tmpDir = tmpdir();
      const filePath = join(tmpDir, `vba-migrate-${Date.now()}-${Math.random().toString(16).slice(2)}.mjs`);
      writeFileSync(filePath, code, "utf8");
      try {
        await execFileAsync("node", ["--check", filePath]);
        return { ok: true };
      } catch (error) {
        return { ok: false, error: error?.stderr?.toString?.() ?? browserSafeErrorMessage(error) };
      }
    }

    // In browser contexts we cannot rely on Node's parser. Instead, strip ESM
    // export syntax (which `new Function` cannot parse) and ask the JS engine to
    // parse the remainder. This catches the common syntax errors without
    // executing the generated code.
    try {
      const script = String(code || "")
        // `export default async function main...` -> `async function main...`
        .replace(/^\s*export\s+default\s+/m, "")
        // `export { foo }` -> removed
        .replace(/^\s*export\s+\{[^}]*\}\s*;?\s*$/gm, "")
        // `export const foo = ...` -> `const foo = ...`
        .replace(/^\s*export\s+(?=(const|let|var|function|async|class)\b)/gm, "");
      // eslint-disable-next-line no-new-func
      new Function(script);
      return { ok: true };
    } catch (error) {
      return { ok: false, error: browserSafeErrorMessage(error) };
    }
  }

  throw new Error(`Unknown target: ${target}`);
}
