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

export const OPEN_FILE_EXTENSIONS = new Set(OPEN_FILE_FILTERS.flatMap((filter) => filter.extensions.map((ext) => ext.toLowerCase())));

export function isOpenWorkbookPath(path: string): boolean {
  const raw = String(path ?? "").trim();
  if (!raw) return false;

  const lastSlash = raw.lastIndexOf("/");
  const lastBackslash = raw.lastIndexOf("\\");
  const lastSep = Math.max(lastSlash, lastBackslash);
  const basename = raw.slice(lastSep + 1);

  const lastDot = basename.lastIndexOf(".");
  if (lastDot <= 0) return false; // no extension or hidden file like `.foo`
  if (lastDot >= basename.length - 1) return false; // empty extension (e.g. `foo.`)

  const ext = basename.slice(lastDot + 1).toLowerCase();
  return OPEN_FILE_EXTENSIONS.has(ext);
}
