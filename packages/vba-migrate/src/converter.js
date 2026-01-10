import { analyzeVbaModule } from "./analyzer.js";
import { postProcessGeneratedCode, validateGeneratedCodeCompiles } from "./postprocess.js";

function buildPrompt({ module, analysis, target }) {
  const language = target === "python" ? "Python" : "TypeScript";
  const apiHint =
    target === "python"
      ? [
          "Use the Formula Python API:",
          "- import formula",
          "- sheet = formula.active_sheet",
          '- write values with: sheet["A1"] = 123',
          '- write formulas with: sheet["A1"].formula = "=A1+A2"',
          "",
          "Return ONLY code (no markdown commentary)."
        ].join("\n")
      : [
          "Use the Formula scripting TypeScript API:",
          "- export default async function main(ctx) { ... }",
          "- const sheet = ctx.activeSheet",
          '- write values with: sheet.range("A1").value = 123',
          '- write formulas with: sheet.range("A1").formula = "=A1+A2"',
          "",
          "Return ONLY code (no markdown commentary)."
        ].join("\n");

  const analysisHints = [
    `Object model usage: Range=${analysis.objectModelUsage.Range.length}, Cells=${analysis.objectModelUsage.Cells.length}, Worksheets=${analysis.objectModelUsage.Worksheets.length}`,
    `External references: ${analysis.externalReferences.length}`,
    `Unsupported constructs: ${analysis.unsupportedConstructs.length}`
  ].join("\n");

  return [
    `Convert the following VBA module to ${language}.`,
    apiHint,
    "",
    "Migration notes:",
    analysisHints,
    "",
    "VBA module:",
    `--- BEGIN VBA (${module.name}) ---`,
    module.code,
    `--- END VBA (${module.name}) ---`
  ].join("\n");
}

export class VbaMigrator {
  constructor({ llm }) {
    if (!llm) throw new Error("VbaMigrator requires an llm client");
    this.llm = llm;
  }

  analyzeModule(module) {
    return analyzeVbaModule(module);
  }

  async convertModule(module, { target }) {
    const analysis = analyzeVbaModule(module);
    const prompt = buildPrompt({ module, analysis, target });

    const raw = await this.llm.complete({
      prompt,
      temperature: 0.0
    });

    const postProcessed = await postProcessGeneratedCode({ code: raw, target });
    const compileCheck = await validateGeneratedCodeCompiles({ code: postProcessed, target });
    if (!compileCheck.ok) {
      throw new Error(`Generated ${target} did not compile: ${compileCheck.error}`);
    }

    return {
      target,
      analysis,
      prompt,
      raw,
      code: postProcessed
    };
  }
}

