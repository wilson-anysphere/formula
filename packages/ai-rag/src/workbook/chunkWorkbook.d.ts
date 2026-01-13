export function chunkWorkbook(
  workbook: any,
  options?: {
    signal?: AbortSignal;
    extractMaxRows?: number;
    extractMaxCols?: number;
    detectRegionsCellLimit?: number;
    maxDataRegionsPerSheet?: number;
    maxFormulaRegionsPerSheet?: number;
    maxRegionsPerSheet?: number;
  }
): any[];
