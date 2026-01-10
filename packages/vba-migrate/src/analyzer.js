function addFinding(targetArray, { line, text, match }) {
  targetArray.push({
    line,
    text: text.trim(),
    match
  });
}

export function analyzeVbaModule(module) {
  const name = module?.name ?? "Unknown";
  const code = module?.code ?? "";
  const lines = String(code).split(/\r?\n/);

  const objectModelUsage = {
    Range: [],
    Cells: [],
    Worksheets: []
  };

  const externalReferences = [];
  const unsupportedConstructs = [];
  const warnings = [];
  const todos = [];

  for (let i = 0; i < lines.length; i += 1) {
    const lineNumber = i + 1;
    const rawLine = lines[i];
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith("'")) continue;
    if (/^\s*Rem\b/i.test(rawLine)) continue;

    if (/\bRange\s*\(/i.test(line) || /\.Range\s*\(/i.test(line)) {
      addFinding(objectModelUsage.Range, { line: lineNumber, text: rawLine, match: "Range" });
    }

    if (/\bCells\s*\(/i.test(line) || /\.Cells\s*\(/i.test(line)) {
      addFinding(objectModelUsage.Cells, { line: lineNumber, text: rawLine, match: "Cells" });
    }

    if (/\bWorksheets\s*\(/i.test(line) || /\bSheets\s*\(/i.test(line)) {
      addFinding(objectModelUsage.Worksheets, { line: lineNumber, text: rawLine, match: "Worksheets" });
    }

    // External references: Windows API declares, COM automation, Shell calls, etc.
    if (/^\s*Declare\b/i.test(rawLine) && /\bLib\s+"[^"]+"/i.test(rawLine)) {
      addFinding(externalReferences, { line: lineNumber, text: rawLine, match: "Declare Lib" });
      todos.push({
        line: lineNumber,
        message: "External Declare Lib call requires a manual replacement in the target runtime."
      });
    }

    if (/\bCreateObject\s*\(/i.test(line) || /\bGetObject\s*\(/i.test(line)) {
      addFinding(externalReferences, { line: lineNumber, text: rawLine, match: "CreateObject/GetObject" });
      todos.push({
        line: lineNumber,
        message: "COM automation object creation has no direct equivalent; migrate to a native library/API."
      });
    }

    if (/\bShell\s*\(/i.test(line)) {
      addFinding(externalReferences, { line: lineNumber, text: rawLine, match: "Shell" });
      todos.push({
        line: lineNumber,
        message: "Shell execution should be migrated to an explicit, permission-checked API (or removed)."
      });
    }

    // Unsupported / risky constructs for an initial migrator
    if (/^\s*On\s+Error\b/i.test(rawLine)) {
      addFinding(unsupportedConstructs, { line: lineNumber, text: rawLine, match: "On Error" });
      warnings.push({
        line: lineNumber,
        message: "VBA error handling (On Error ...) does not map cleanly; review translated try/except."
      });
    }

    if (/\bGoTo\b/i.test(line) || /\bGoSub\b/i.test(line)) {
      addFinding(unsupportedConstructs, { line: lineNumber, text: rawLine, match: "GoTo/GoSub" });
      warnings.push({
        line: lineNumber,
        message: "GoTo/GoSub control flow typically needs refactoring into loops/functions."
      });
    }

    if (/\bWithEvents\b/i.test(line) || /\bRaiseEvent\b/i.test(line)) {
      addFinding(unsupportedConstructs, { line: lineNumber, text: rawLine, match: "Events" });
      warnings.push({
        line: lineNumber,
        message: "VBA events do not directly translate; consider explicit event hooks in the target platform."
      });
    }

    if (/\bDoEvents\b/i.test(line)) {
      addFinding(unsupportedConstructs, { line: lineNumber, text: rawLine, match: "DoEvents" });
      warnings.push({
        line: lineNumber,
        message: "DoEvents is UI-thread specific; translated scripts should be async/await friendly."
      });
    }
  }

  // De-dupe TODOs / warnings by line+message to keep report stable.
  const stableKey = (item) => `${item.line}:${item.message}`;
  const uniq = (items) => {
    const seen = new Set();
    const out = [];
    for (const item of items) {
      const key = stableKey(item);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(item);
    }
    return out;
  };

  return {
    moduleName: name,
    objectModelUsage,
    externalReferences,
    unsupportedConstructs,
    warnings: uniq(warnings),
    todos: uniq(todos)
  };
}

export function migrationReportToMarkdown(report) {
  const usageCount = (arr) => arr?.length ?? 0;
  const lines = [];
  lines.push(`# VBA Migration Report: ${report.moduleName}`);
  lines.push("");
  lines.push("## Object model usage");
  lines.push(`- Range: ${usageCount(report.objectModelUsage?.Range)}`);
  lines.push(`- Cells: ${usageCount(report.objectModelUsage?.Cells)}`);
  lines.push(`- Worksheets/Sheets: ${usageCount(report.objectModelUsage?.Worksheets)}`);
  lines.push("");

  if (report.externalReferences?.length) {
    lines.push("## External references");
    for (const ref of report.externalReferences) {
      lines.push(`- L${ref.line}: ${ref.text.trim()}`);
    }
    lines.push("");
  }

  if (report.unsupportedConstructs?.length) {
    lines.push("## Unsupported / risky constructs");
    for (const item of report.unsupportedConstructs) {
      lines.push(`- L${item.line}: ${item.text.trim()}`);
    }
    lines.push("");
  }

  if (report.warnings?.length) {
    lines.push("## Warnings");
    for (const warning of report.warnings) {
      lines.push(`- L${warning.line}: ${warning.message}`);
    }
    lines.push("");
  }

  if (report.todos?.length) {
    lines.push("## TODOs");
    for (const todo of report.todos) {
      lines.push(`- L${todo.line}: ${todo.message}`);
    }
    lines.push("");
  }

  return lines.join("\n");
}

