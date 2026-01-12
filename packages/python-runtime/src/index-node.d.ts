export class NativePythonRuntime {
  constructor(options?: any);
  initialize?(options?: any): Promise<void>;
  execute?(code: string, options?: any): Promise<any>;
  destroy?(): void;
}

export class PyodideRuntime {
  initialized: boolean;
  constructor(options?: any);
  getBackendMode(): string;
  initialize(options?: any): Promise<void>;
  execute(code: string, options?: any): Promise<any>;
  destroy(): void;
}

export const formulaFiles: any;

export const __pyodideMocks: any;
