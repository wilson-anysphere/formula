import type { DocumentController } from "../document/documentController.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { applyNumberFormatPreset } from "../formatting/toolbar.js";
import type { CellRange } from "../formatting/toolbar.js";

export type ApplyFormattingToSelection = (
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
  options?: { forceBatch?: boolean },
) => void;

function parseDecimalPlaces(format: string): number {
  const dot = format.indexOf(".");
  if (dot === -1) return 0;
  let count = 0;
  for (let i = dot + 1; i < format.length; i += 1) {
    const ch = format[i];
    if (ch === "0" || ch === "#") count += 1;
    else break;
  }
  return count;
}

function stepDecimalPlacesInNumberFormat(
  format: string | null,
  direction: "increase" | "decrease",
): string | null {
  const raw = (format ?? "").trim();
  const section = (raw.split(";")[0] ?? "").trim();
  const lower = section.toLowerCase();
  const compact = lower.replace(/\s+/g, "");
  // Avoid trying to manipulate date/time format codes.
  if (compact.includes("m/d/yyyy") || compact.includes("yyyy-mm-dd")) return null;
  if (/^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact)) return null;

  const currencyMatch = /[$€£¥]/.exec(section);
  const prefix = currencyMatch?.[0] ?? "";
  const suffix = section.includes("%") ? "%" : "";
  const useThousands = section.includes(",");
  const decimals = parseDecimalPlaces(section);

  const nextDecimals =
    direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
  if (nextDecimals === decimals) return null;

  const integer = useThousands ? "#,##0" : "0";
  const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
  return `${prefix}${integer}${fraction}${suffix}`;
}

function applyConstantNumberFormatPatch(
  patch: { numberFormat: string | null },
): (doc: DocumentController, sheetId: string, ranges: CellRange[]) => boolean {
  return (doc, sheetId, ranges) => {
    let applied = true;
    for (const range of ranges) {
      const ok = doc.setRangeFormat(sheetId, range, patch, { label: "Number format" });
      if (ok === false) applied = false;
    }
    return applied;
  };
}

export function registerNumberFormatCommands(params: {
  commandRegistry: CommandRegistry;
  applyFormattingToSelection: ApplyFormattingToSelection;
  getActiveCellNumberFormat: () => string | null;
  t: (key: string) => string;
  category?: string | null;
}): void {
  const { commandRegistry, applyFormattingToSelection, getActiveCellNumberFormat, t } = params;
  const category = params.category ?? null;

  const register = (commandId: string, titleKey: string, run: () => void) => {
    commandRegistry.registerBuiltinCommand(commandId, t(titleKey), run, { category });
  };

  register("format.numberFormat.general", "command.format.numberFormat.general", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.general"),
      applyConstantNumberFormatPatch({ numberFormat: null }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.number", "command.format.numberFormat.number", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.number"),
      applyConstantNumberFormatPatch({ numberFormat: "0.00" }),
      { forceBatch: true },
    ),
  );

  // Existing canonical commands (used by context menus + existing shortcuts).
  register("format.numberFormat.currency", "command.format.numberFormat.currency", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.currency"),
      (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.accounting", "command.format.numberFormat.accounting", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.accounting"),
      (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"),
      { forceBatch: true },
    ),
  );

  // Back-compat alias: existing keybindings and context menus still reference `format.numberFormat.date`.
  register("format.numberFormat.date", "command.format.numberFormat.date", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.date"),
      (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.shortDate", "command.format.numberFormat.shortDate", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.shortDate"),
      (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.longDate", "command.format.numberFormat.longDate", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.longDate"),
      applyConstantNumberFormatPatch({ numberFormat: "yyyy-mm-dd" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.time", "command.format.numberFormat.time", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.time"),
      applyConstantNumberFormatPatch({ numberFormat: "h:mm:ss" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.percent", "command.format.numberFormat.percent", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.percent"),
      (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.fraction", "command.format.numberFormat.fraction", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.fraction"),
      applyConstantNumberFormatPatch({ numberFormat: "# ?/?" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.scientific", "command.format.numberFormat.scientific", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.scientific"),
      applyConstantNumberFormatPatch({ numberFormat: "0.00E+00" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.text", "command.format.numberFormat.text", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.text"),
      applyConstantNumberFormatPatch({ numberFormat: "@" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.commaStyle", "command.format.numberFormat.commaStyle", () =>
    applyFormattingToSelection(
      t("command.format.numberFormat.commaStyle"),
      applyConstantNumberFormatPatch({ numberFormat: "#,##0.00" }),
      { forceBatch: true },
    ),
  );

  register("format.numberFormat.increaseDecimal", "command.format.numberFormat.increaseDecimal", () => {
    const next = stepDecimalPlacesInNumberFormat(getActiveCellNumberFormat(), "increase");
    if (!next) return;
    applyFormattingToSelection(
      t("command.format.numberFormat.increaseDecimal"),
      applyConstantNumberFormatPatch({ numberFormat: next }),
      { forceBatch: true },
    );
  });

  register("format.numberFormat.decreaseDecimal", "command.format.numberFormat.decreaseDecimal", () => {
    const next = stepDecimalPlacesInNumberFormat(getActiveCellNumberFormat(), "decrease");
    if (!next) return;
    applyFormattingToSelection(
      t("command.format.numberFormat.decreaseDecimal"),
      applyConstantNumberFormatPatch({ numberFormat: next }),
      { forceBatch: true },
    );
  });

  // Accounting symbol picker menu items (used by the ribbon's accounting dropdown).
  const accountingSymbols: Array<{ id: string; titleKey: string; symbol: string }> = [
    { id: "usd", titleKey: "command.format.numberFormat.accounting.usd", symbol: "$" },
    { id: "eur", titleKey: "command.format.numberFormat.accounting.eur", symbol: "€" },
    { id: "gbp", titleKey: "command.format.numberFormat.accounting.gbp", symbol: "£" },
    { id: "jpy", titleKey: "command.format.numberFormat.accounting.jpy", symbol: "¥" },
  ];
  for (const { id, titleKey, symbol } of accountingSymbols) {
    register(`format.numberFormat.accounting.${id}`, titleKey, () =>
      applyFormattingToSelection(
        t(titleKey),
        applyConstantNumberFormatPatch({ numberFormat: `${symbol}#,##0.00` }),
        { forceBatch: true },
      ),
    );
  }
}
