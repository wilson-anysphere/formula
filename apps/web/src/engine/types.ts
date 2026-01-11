export type EngineCellScalar = number | string | boolean | null;

export type EngineWorkbookJson = {
  sheets: Record<string, { cells: Record<string, EngineCellScalar> }>;
};

