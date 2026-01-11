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
    Worksheets: [],
    Workbook: [],
    Application: [],
    ActiveSheet: [],
    ActiveCell: [],
    Selection: []
  };

  const rangeShapes = {
    singleCell: [],
    multiCell: [],
    rows: [],
    columns: [],
    other: []
  };

  const externalReferences = [];
  const unsafeConstructs = [];
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
      const rangeArg = /\bRange\s*\(\s*"(?<addr>[^"]+)"\s*\)/i.exec(rawLine)?.groups?.addr;
      if (rangeArg) {
        const addr = rangeArg.trim();
        if (/^[A-Za-z]+\d+$/.test(addr)) {
          addFinding(rangeShapes.singleCell, { line: lineNumber, text: rawLine, match: addr });
        } else if (/^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(addr)) {
          addFinding(rangeShapes.multiCell, { line: lineNumber, text: rawLine, match: addr });
        } else if (/^\d+:\d+$/.test(addr)) {
          addFinding(rangeShapes.rows, { line: lineNumber, text: rawLine, match: addr });
        } else if (/^[A-Za-z]+:[A-Za-z]+$/.test(addr)) {
          addFinding(rangeShapes.columns, { line: lineNumber, text: rawLine, match: addr });
        } else {
          addFinding(rangeShapes.other, { line: lineNumber, text: rawLine, match: addr });
        }
      }
    }

    if (/\bCells\s*\(/i.test(line) || /\.Cells\s*\(/i.test(line)) {
      addFinding(objectModelUsage.Cells, { line: lineNumber, text: rawLine, match: "Cells" });
    }

    if (/\bWorksheets\s*\(/i.test(line) || /\bSheets\s*\(/i.test(line)) {
      addFinding(objectModelUsage.Worksheets, { line: lineNumber, text: rawLine, match: "Worksheets" });
    }

    if (/\b(ThisWorkbook|ActiveWorkbook|Workbooks)\b/i.test(line)) {
      addFinding(objectModelUsage.Workbook, { line: lineNumber, text: rawLine, match: "Workbook" });
    }

    if (/\bApplication\b/i.test(line)) {
      addFinding(objectModelUsage.Application, { line: lineNumber, text: rawLine, match: "Application" });
    }

    if (/\bActiveSheet\b/i.test(line)) {
      addFinding(objectModelUsage.ActiveSheet, { line: lineNumber, text: rawLine, match: "ActiveSheet" });
    }

    if (/\bActiveCell\b/i.test(line)) {
      addFinding(objectModelUsage.ActiveCell, { line: lineNumber, text: rawLine, match: "ActiveCell" });
    }

    if (/\bSelection\b/i.test(line)) {
      addFinding(objectModelUsage.Selection, { line: lineNumber, text: rawLine, match: "Selection" });
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

    if (/\bFileSystemObject\b/i.test(line) || /\bScripting\.FileSystemObject\b/i.test(line)) {
      addFinding(externalReferences, { line: lineNumber, text: rawLine, match: "FileSystemObject" });
      todos.push({
        line: lineNumber,
        message: "File system access (FileSystemObject) must be migrated to a permission-checked API."
      });
    }

    // Unsafe dynamic execution patterns.
    if (/\bEvaluate\s*\(/i.test(line) || /\bApplication\.Evaluate\b/i.test(line)) {
      addFinding(unsafeConstructs, { line: lineNumber, text: rawLine, match: "Evaluate" });
      warnings.push({
        line: lineNumber,
        message: "Evaluate() executes formula strings dynamically; translated code should avoid eval-like behavior."
      });
    }

    // VBA `Execute` executes a string as code.
    if (/\bExecute\b/i.test(line)) {
      addFinding(unsafeConstructs, { line: lineNumber, text: rawLine, match: "Execute" });
      warnings.push({
        line: lineNumber,
        message: "Execute executes arbitrary VBA source; translated code must replace it with explicit logic."
      });
    }

    if (/\bCallByName\b/i.test(line)) {
      addFinding(unsafeConstructs, { line: lineNumber, text: rawLine, match: "CallByName" });
      warnings.push({
        line: lineNumber,
        message: "CallByName performs dynamic dispatch; translated code should use direct calls and validate inputs."
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

  const uniqRiskFactors = (items) => {
    const seen = new Set();
    const out = [];
    for (const item of items) {
      const key = `${item.code}:${item.line}:${item.message}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(item);
    }
    return out;
  };

  const riskFactors = [];
  for (const ref of externalReferences) {
    riskFactors.push({
      code: "external_dependency",
      line: ref.line,
      message: `External dependency detected (${ref.match}).`
    });
  }
  for (const item of unsafeConstructs) {
    riskFactors.push({
      code: "unsafe_dynamic_execution",
      line: item.line,
      message: `Unsafe dynamic execution detected (${item.match}).`
    });
  }
  for (const item of unsupportedConstructs) {
    riskFactors.push({
      code: "unsupported_construct",
      line: item.line,
      message: `Unsupported/risky construct detected (${item.match}).`
    });
  }

  const riskScore = Math.min(
    100,
    externalReferences.length * 25 + unsafeConstructs.length * 30 + unsupportedConstructs.length * 10
  );
  const riskLevel = riskScore >= 70 ? "high" : riskScore >= 30 ? "medium" : "low";

  return {
    moduleName: name,
    objectModelUsage,
    rangeShapes,
    externalReferences,
    unsafeConstructs,
    unsupportedConstructs,
    warnings: uniq(warnings),
    todos: uniq(todos),
    risk: {
      score: riskScore,
      level: riskLevel,
      factors: uniqRiskFactors(riskFactors)
    }
  };
}

export function migrationReportToMarkdown(report) {
  const usageCount = (arr) => arr?.length ?? 0;
  const lines = [];
  lines.push(`# VBA Migration Report: ${report.moduleName}`);
  lines.push("");
  if (report.risk) {
    lines.push("## Risk score");
    lines.push(`- Score: ${report.risk.score} (${report.risk.level})`);
    lines.push("");
  }
  lines.push("## Object model usage");
  lines.push(`- Range: ${usageCount(report.objectModelUsage?.Range)}`);
  lines.push(`- Cells: ${usageCount(report.objectModelUsage?.Cells)}`);
  lines.push(`- Worksheets/Sheets: ${usageCount(report.objectModelUsage?.Worksheets)}`);
  lines.push(`- Workbook: ${usageCount(report.objectModelUsage?.Workbook)}`);
  lines.push(`- Application: ${usageCount(report.objectModelUsage?.Application)}`);
  lines.push(`- ActiveSheet: ${usageCount(report.objectModelUsage?.ActiveSheet)}`);
  lines.push(`- ActiveCell: ${usageCount(report.objectModelUsage?.ActiveCell)}`);
  lines.push(`- Selection: ${usageCount(report.objectModelUsage?.Selection)}`);
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

  if (report.unsafeConstructs?.length) {
    lines.push("## Unsafe dynamic execution");
    for (const item of report.unsafeConstructs) {
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
