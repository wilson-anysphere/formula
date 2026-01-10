import { writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { rowColToA1 } from "./a1.js";

const execFileAsync = promisify(execFile);

function stripMarkdownCodeFences(text) {
  const trimmed = String(text || "").trim();
  const fenceMatch = /^```[a-zA-Z0-9_-]*\n([\s\S]*?)\n```$/m.exec(trimmed);
  if (!fenceMatch) return trimmed;
  return fenceMatch[1].trim();
}

function ensurePythonWrapper(code) {
  const cleaned = String(code || "").trim();
  const hasMain = /\bdef\s+main\s*\(/.test(cleaned);
  if (hasMain) return cleaned;

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
  out = out.replace(/\bRange\(/g, "range(");
  out = out.replace(/\bCells\(/g, "cell(");
  out = out.replace(/\.Value\b/g, ".value");
  out = out.replace(/\.Formula\b/g, ".formula");
  return out;
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
    const tmpDir = os.tmpdir();
    const filePath = path.join(tmpDir, `vba-migrate-${Date.now()}-${Math.random().toString(16).slice(2)}.py`);
    writeFileSync(filePath, code, "utf8");
    try {
      await execFileAsync("python", ["-m", "py_compile", filePath]);
      return { ok: true };
    } catch (error) {
      return { ok: false, error: error?.stderr?.toString?.() ?? error?.message ?? String(error) };
    }
  }

  if (target === "typescript") {
    // We intentionally restrict generated code to TS that is also valid JS/ESM.
    // Use Node's parser (`node --check`) to validate module syntax deterministically.
    const tmpDir = os.tmpdir();
    const filePath = path.join(tmpDir, `vba-migrate-${Date.now()}-${Math.random().toString(16).slice(2)}.mjs`);
    writeFileSync(filePath, code, "utf8");
    try {
      await execFileAsync("node", ["--check", filePath]);
      return { ok: true };
    } catch (error) {
      return { ok: false, error: error?.stderr?.toString?.() ?? error?.message ?? String(error) };
    }
  }

  throw new Error(`Unknown target: ${target}`);
}
