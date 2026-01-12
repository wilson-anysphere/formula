import { t } from "./i18n/index.js";

export function getOpenFileFilters() {
  return [
    {
      name: t("fileDialog.filters.excelWorkbook"),
      extensions: [
        "xlsx",
        "xlsm",
        "xltx",
        "xltm",
        "xlam",
        "xls",
        "xlt",
        "xla",
        "csv",
        "parquet",
      ],
    },
    {
      name: t("fileDialog.filters.excelBinaryWorkbook"),
      extensions: ["xlsb"],
    },
  ];
}

export const OPEN_FILE_FILTERS = getOpenFileFilters();
