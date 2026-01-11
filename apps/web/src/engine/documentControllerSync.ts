import type { EngineWorkbookJson } from "./types";

export const DEMO_WORKBOOK: EngineWorkbookJson = {
  sheets: {
    Sheet1: {
      cells: {
        A1: 1,
        A2: 2,
        B1: "=A1+A2",
        B2: "=B1*2",
        C1: "hello",
      },
    },
  },
};

export const DEMO_WORKBOOK_JSON = JSON.stringify(DEMO_WORKBOOK);
