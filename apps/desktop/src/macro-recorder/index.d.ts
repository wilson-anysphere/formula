export class MacroRecorder {
  constructor(workbook: any);
  readonly recording: boolean;

  start(): void;
  stop(): void;
  clear(): void;

  getRawActions(): any[];
  getOptimizedActions(): any[];
}

export function optimizeMacroActions(actions: any[]): any[];
export function generatePythonMacro(actions: any[]): string;
export function generateTypeScriptMacro(actions: any[]): string;

