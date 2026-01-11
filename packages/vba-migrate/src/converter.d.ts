import type { VbaModule } from "./analyzer.js";

export class VbaMigrator {
  constructor(params: { llm: any });
  llm: any;

  completePrompt(prompt: string, options?: { temperature?: number }): Promise<string>;
  analyzeModule(module: VbaModule): any;
  convertModule(module: VbaModule, options: { target: "python" | "typescript" }): Promise<any>;
}

