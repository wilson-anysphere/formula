import type { MacroBackend, MacroInfo, MacroRunResult } from "./types";
import type { MacroSecurityController } from "./security";

export interface RunMacroOptions {
  workbookId: string;
  macroId: string;
  timeoutMs?: number;
}

export class MacroRunner {
  constructor(
    private readonly backend: MacroBackend,
    private readonly security: MacroSecurityController
  ) {}

  async list(workbookId: string): Promise<MacroInfo[]> {
    return await this.backend.listMacros(workbookId);
  }

  async run(options: RunMacroOptions): Promise<MacroRunResult> {
    const setting = await this.security.getSetting(options.workbookId);

    if (setting === "disabled") {
      return {
        ok: false,
        output: [],
        error: { message: "Macros are disabled for this workbook." },
      };
    }

    let decision = { enabled: true, permissions: [] as const };
    if (setting === "prompt") {
      decision = await this.security.requestEnableMacros(options.workbookId);
      if (!decision.enabled) {
        return {
          ok: false,
          output: [],
          error: { message: "User declined to enable macros." },
        };
      }
    }

    return await this.backend.runMacro({
      workbookId: options.workbookId,
      macroId: options.macroId,
      permissions: [...decision.permissions],
      timeoutMs: options.timeoutMs,
    });
  }
}

