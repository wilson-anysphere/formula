#!/usr/bin/env node
import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { analyzeVbaModule } from "./analyzer.js";
import { VbaMigrator } from "./converter.js";
import { rowColToA1 } from "./a1.js";
import { Workbook } from "./workbook.js";
import { validateMigration } from "./validator.js";
import { RustCliOracle } from "./vba/oracle.js";

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      args._.push(token);
      continue;
    }
    const key = token.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    i += 1;
  }
  return args;
}

async function findFiles(dir, { exts }) {
  const out = [];
  const entries = await (await import("node:fs/promises")).readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await findFiles(full, { exts })));
      continue;
    }
    const ext = path.extname(entry.name).toLowerCase();
    if (exts.includes(ext)) out.push(full);
  }
  out.sort((a, b) => a.localeCompare(b));
  return out;
}

function extractFirstSubName(vbaSource) {
  const match = /^\s*Sub\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\b/im.exec(String(vbaSource || ""));
  return match?.groups?.name ?? null;
}

class RuleBasedLlmClient {
  async complete({ prompt }) {
    const isPython = /to Python/i.test(prompt);
    const moduleMatch = /--- BEGIN VBA \((?<name>[^)]+)\) ---\n(?<code>[\s\S]*?)\n--- END VBA/m.exec(prompt);
    const code = moduleMatch?.groups?.code ?? "";

    // Extremely small subset for batch validation: literal Range/Cells assignments.
    const lines = code.split(/\r?\n/);
    const body = [];
    for (const raw of lines) {
      const line = raw.trim();
      if (!line) continue;
      if (line.startsWith("'")) continue;
      if (/^\s*Rem\b/i.test(raw)) continue;

      const rangeValue = /\bRange\(\s*"(?<addr>[^"]+)"\s*\)\.Value\s*=\s*(?<expr>.+)$/i.exec(raw);
      if (rangeValue) {
        body.push({ kind: "value", addr: rangeValue.groups.addr, expr: rangeValue.groups.expr.trim() });
        continue;
      }
      const rangeFormula = /\bRange\(\s*"(?<addr>[^"]+)"\s*\)\.Formula\s*=\s*(?<expr>.+)$/i.exec(raw);
      if (rangeFormula) {
        body.push({ kind: "formula", addr: rangeFormula.groups.addr, expr: rangeFormula.groups.expr.trim() });
        continue;
      }
      const cellsValue = /\bCells\(\s*(?<row>\d+)\s*,\s*(?<col>\d+)\s*\)\.Value\s*=\s*(?<expr>.+)$/i.exec(raw);
      if (cellsValue) {
        body.push({
          kind: "cellValue",
          row: Number(cellsValue.groups.row),
          col: Number(cellsValue.groups.col),
          expr: cellsValue.groups.expr.trim(),
        });
      }
    }

    const parseLiteral = (expr, { python } = {}) => {
      const trimmed = expr.trim();
      const str = /^"(?<s>(?:[^"]|"")*)"$/.exec(trimmed)?.groups?.s;
      if (str !== undefined) return JSON.stringify(str.replace(/""/g, '"'));
      if (/^(True|False)$/i.test(trimmed)) {
        if (python) return /^True$/i.test(trimmed) ? "True" : "False";
        return /^True$/i.test(trimmed) ? "true" : "false";
      }
      if (/^[+-]?\d+(\.\d+)?$/.test(trimmed)) return trimmed;
      // Fallback: emit as string
      return JSON.stringify(trimmed);
    };

    if (isPython) {
      const out = ["sheet = formula.active_sheet"];
      for (const stmt of body) {
        if (stmt.kind === "value") out.push(`sheet["${stmt.addr}"] = ${parseLiteral(stmt.expr, { python: true })}`);
        if (stmt.kind === "formula")
          out.push(`sheet["${stmt.addr}"].formula = ${parseLiteral(stmt.expr, { python: true })}`);
        if (stmt.kind === "cellValue") {
          const addr = rowColToA1(stmt.row, stmt.col);
          out.push(`sheet["${addr}"] = ${parseLiteral(stmt.expr, { python: true })}`);
        }
      }
      return out.join("\n");
    }

    const out = ["const sheet = ctx.activeSheet;"];
    for (const stmt of body) {
      if (stmt.kind === "value") out.push(`sheet.range("${stmt.addr}").value = ${parseLiteral(stmt.expr)};`);
      if (stmt.kind === "formula") out.push(`sheet.range("${stmt.addr}").formula = ${parseLiteral(stmt.expr)};`);
      if (stmt.kind === "cellValue")
        out.push(`sheet.cell(${stmt.row}, ${stmt.col}).value = ${parseLiteral(stmt.expr)};`);
    }
    return out.join("\n");
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const dir = args.dir ?? args._[0];
  if (!dir) {
    console.error("Usage: vba-migrate --dir <path> [--target python|typescript|both] [--entryPoint Main]");
    process.exit(2);
  }

  const target = args.target ?? "both";
  const entryPointOverride = args.entryPoint ?? null;

  const oracle = new RustCliOracle();
  const llm = new RuleBasedLlmClient();
  const migrator = new VbaMigrator({ llm });

  const files = await findFiles(path.resolve(dir), { exts: [".xlsm", ".json"] });
  const results = [];

  for (const file of files) {
    const bytes = await readFile(file);
    const ext = path.extname(file).toLowerCase();

    let workbookBytes;
    let vbaModules;
    let procedures;

    if (ext === ".xlsm") {
      const extracted = await oracle.extract({ workbookBytes: bytes });
      if (!extracted.ok) {
        results.push({
          file,
          ok: false,
          error: extracted.error ?? "Failed to extract workbook",
        });
        continue;
      }
      workbookBytes = extracted.workbookBytes;
      vbaModules = extracted.workbook?.vbaModules ?? [];
      procedures = extracted.procedures ?? [];
    } else {
      workbookBytes = bytes;
      const payload = JSON.parse(bytes.toString("utf8"));
      vbaModules = payload?.vbaModules ?? [];
      procedures = [];
    }

    const workbook = Workbook.fromBytes(workbookBytes);
    const module = vbaModules[0] ?? null;
    if (!module) {
      results.push({ file, ok: false, error: "No VBA modules found" });
      continue;
    }

    const analysis = analyzeVbaModule(module);
    const entryPoint =
      entryPointOverride ?? procedures[0]?.name ?? extractFirstSubName(module.code) ?? "Main";

    const targets = target === "both" ? ["python", "typescript"] : [target];
    const perTarget = [];
    for (const tgt of targets) {
      const conversion = await migrator.convertModule(module, { target: tgt });
      const validation = await validateMigration({
        workbook,
        module,
        entryPoint,
        target: tgt,
        code: conversion.code,
        oracle,
      });
      perTarget.push({
        target: tgt,
        ok: validation.ok,
        mismatches: validation.mismatches,
        vbaDiff: validation.vbaDiff,
        scriptDiff: validation.scriptDiff,
        oracle: validation.oracle,
      });
    }

    results.push({
      file,
      ok: perTarget.every((r) => r.ok),
      entryPoint,
      risk: analysis.risk,
      results: perTarget,
    });
  }

  const summary = {
    ok: results.every((r) => r.ok),
    totals: {
      files: results.length,
      passed: results.filter((r) => r.ok).length,
      failed: results.filter((r) => !r.ok).length,
    },
    results,
  };

  process.stdout.write(JSON.stringify(summary, null, 2) + "\n");
  process.exit(summary.ok ? 0 : 1);
}

main().catch((err) => {
  console.error(err?.stack ?? err?.message ?? String(err));
  process.exit(1);
});
