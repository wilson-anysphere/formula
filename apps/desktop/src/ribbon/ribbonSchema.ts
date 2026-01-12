export type RibbonButtonKind = "button" | "toggle" | "dropdown";
export type RibbonButtonSize = "large" | "small" | "icon";

export interface RibbonMenuItemDefinition {
  /**
   * Stable command identifier (used for wiring actions).
   */
  id: string;
  label: string;
  ariaLabel: string;
  /**
   * Small text glyph used as a placeholder until a real icon system exists.
   */
  icon?: string;
  /**
   * Optional E2E hook.
   */
  testId?: string;
  disabled?: boolean;
}

export interface RibbonButtonDefinition {
  /**
   * Stable command identifier (used for wiring actions).
   *
   * Convention: `{tab}.{group}.{command}` (e.g. `home.clipboard.paste`).
   */
  id: string;
  label: string;
  ariaLabel: string;
  /**
   * Small text glyph used as a placeholder until a real icon system exists.
   */
  icon?: string;
  kind?: RibbonButtonKind;
  size?: RibbonButtonSize;
  /**
   * Optional dropdown menu items. When provided for a `kind: "dropdown"` button,
   * the ribbon will render a menu instead of invoking the command directly.
   */
  menuItems?: RibbonMenuItemDefinition[];
  /**
   * Optional E2E hook.
   */
  testId?: string;
  /**
   * Initial pressed state for toggle buttons (purely UI; can be replaced with
   * app-driven state later).
   */
  defaultPressed?: boolean;
  disabled?: boolean;
}

export interface RibbonGroupDefinition {
  id: string;
  label: string;
  buttons: RibbonButtonDefinition[];
}

export interface RibbonTabDefinition {
  id: string;
  label: string;
  groups: RibbonGroupDefinition[];
  /**
   * File tab is typically styled as a primary pill and may later open a
   * backstage view.
   */
  isFile?: boolean;
}

export interface RibbonSchema {
  tabs: RibbonTabDefinition[];
}

export interface RibbonActions {
  /**
   * Called for any command-like activation (including dropdown buttons).
   */
  onCommand?: (commandId: string) => void;
  /**
   * Called when a toggle button changes state.
   */
  onToggle?: (commandId: string, pressed: boolean) => void;
  /**
   * Called when a tab is selected.
   */
  onTabChange?: (tabId: string) => void;
}

export const defaultRibbonSchema: RibbonSchema = {
  tabs: [
    {
      id: "file",
      label: "File",
      isFile: true,
      groups: [
        {
          id: "file.new",
          label: "New",
          buttons: [
            { id: "file.new.new", label: "New", ariaLabel: "New", icon: "📄", kind: "dropdown", size: "large" },
            { id: "file.new.blankWorkbook", label: "Blank workbook", ariaLabel: "Blank workbook", icon: "⬜" },
            { id: "file.new.templates", label: "Templates", ariaLabel: "Templates", icon: "📑", kind: "dropdown" },
          ],
        },
        {
          id: "file.info",
          label: "Info",
          buttons: [
            { id: "file.info.protectWorkbook", label: "Protect Workbook", ariaLabel: "Protect Workbook", icon: "🔒", kind: "dropdown", size: "large" },
            { id: "file.info.inspectWorkbook", label: "Inspect Workbook", ariaLabel: "Inspect Workbook", icon: "🔍", kind: "dropdown" },
            { id: "file.info.manageWorkbook", label: "Manage Workbook", ariaLabel: "Manage Workbook", icon: "🗂", kind: "dropdown" },
          ],
        },
        {
          id: "file.open",
          label: "Open",
          buttons: [
            { id: "file.open.open", label: "Open", ariaLabel: "Open", icon: "📂", size: "large" },
            { id: "file.open.recent", label: "Recent", ariaLabel: "Recent", icon: "🕘", kind: "dropdown" },
            { id: "file.open.pinned", label: "Pinned", ariaLabel: "Pinned", icon: "📌", kind: "dropdown" },
          ],
        },
        {
          id: "file.save",
          label: "Save",
          buttons: [
            { id: "file.save.save", label: "Save", ariaLabel: "Save", icon: "💾", size: "large", testId: "ribbon-save" },
            { id: "file.save.saveAs", label: "Save As", ariaLabel: "Save As", icon: "📝", kind: "dropdown" },
            { id: "file.save.autoSave", label: "AutoSave", ariaLabel: "AutoSave", icon: "⏱", kind: "toggle", defaultPressed: false },
          ],
        },
        {
          id: "file.export",
          label: "Export",
          buttons: [
            { id: "file.export.export", label: "Export", ariaLabel: "Export", icon: "📤", kind: "dropdown", size: "large" },
            { id: "file.export.createPdf", label: "Create PDF/XPS", ariaLabel: "Create PDF or XPS", icon: "📄" },
            { id: "file.export.changeFileType", label: "Change File Type", ariaLabel: "Change File Type", icon: "🔁", kind: "dropdown" },
          ],
        },
        {
          id: "file.print",
          label: "Print",
          buttons: [
            { id: "file.print.print", label: "Print", ariaLabel: "Print", icon: "🖨", size: "large", testId: "ribbon-print" },
            { id: "file.print.printPreview", label: "Print Preview", ariaLabel: "Print Preview", icon: "👁" },
            { id: "file.print.pageSetup", label: "Page Setup", ariaLabel: "Page Setup", icon: "📐", kind: "dropdown" },
          ],
        },
        {
          id: "file.share",
          label: "Share",
          buttons: [
            { id: "file.share.share", label: "Share", ariaLabel: "Share", icon: "🔗", size: "large" },
            { id: "file.share.email", label: "Email", ariaLabel: "Email", icon: "✉️", kind: "dropdown" },
            { id: "file.share.presentOnline", label: "Present Online", ariaLabel: "Present Online", icon: "🌐" },
          ],
        },
        {
          id: "file.options",
          label: "Options",
          buttons: [
            { id: "file.options.options", label: "Options", ariaLabel: "Options", icon: "⚙️", size: "large" },
            { id: "file.options.account", label: "Account", ariaLabel: "Account", icon: "👤" },
            { id: "file.options.close", label: "Close", ariaLabel: "Close", icon: "❌", testId: "ribbon-close" },
          ],
        },
      ],
    },
    {
      id: "home",
      label: "Home",
      groups: [
        {
          id: "home.clipboard",
          label: "Clipboard",
          buttons: [
            {
              id: "home.clipboard.paste",
              label: "Paste",
              ariaLabel: "Paste",
              icon: "📋",
              kind: "dropdown",
              size: "large",
              testId: "ribbon-paste",
              menuItems: [
                { id: "home.clipboard.paste.default", label: "Paste", ariaLabel: "Paste", icon: "📋" },
                { id: "home.clipboard.paste.values", label: "Paste Values", ariaLabel: "Paste Values", icon: "123" },
                { id: "home.clipboard.paste.formulas", label: "Paste Formulas", ariaLabel: "Paste Formulas", icon: "fx" },
                { id: "home.clipboard.paste.formats", label: "Paste Formatting", ariaLabel: "Paste Formatting", icon: "🎨" },
                { id: "home.clipboard.paste.transpose", label: "Transpose", ariaLabel: "Transpose", icon: "🔁" },
              ],
            },
            { id: "home.clipboard.pasteSpecial", label: "Paste Special", ariaLabel: "Paste Special", icon: "📌", kind: "dropdown", size: "small" },
            { id: "home.clipboard.cut", label: "Cut", ariaLabel: "Cut", icon: "✂️", size: "icon" },
            { id: "home.clipboard.copy", label: "Copy", ariaLabel: "Copy", icon: "📄", size: "icon" },
            { id: "home.clipboard.formatPainter", label: "Format Painter", ariaLabel: "Format Painter", icon: "🖌", size: "small" },
            { id: "home.clipboard.clipboardPane", label: "Clipboard", ariaLabel: "Open Clipboard", icon: "📎", kind: "dropdown", size: "small" },
          ],
        },
        {
          id: "home.font",
          label: "Font",
          buttons: [
            {
              id: "home.font.fontName",
              label: "Font",
              ariaLabel: "Font",
              icon: "A",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.font.fontName.calibri", label: "Calibri", ariaLabel: "Calibri", icon: "A" },
                { id: "home.font.fontName.arial", label: "Arial", ariaLabel: "Arial", icon: "A" },
                { id: "home.font.fontName.times", label: "Times New Roman", ariaLabel: "Times New Roman", icon: "A" },
                { id: "home.font.fontName.courier", label: "Courier New", ariaLabel: "Courier New", icon: "A" },
              ],
            },
            {
              id: "home.font.fontSize",
              label: "Size",
              ariaLabel: "Font Size",
              icon: "↕",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.font.fontSize.8", label: "8", ariaLabel: "Font size 8", icon: "8" },
                { id: "home.font.fontSize.9", label: "9", ariaLabel: "Font size 9", icon: "9" },
                { id: "home.font.fontSize.10", label: "10", ariaLabel: "Font size 10", icon: "10" },
                { id: "home.font.fontSize.11", label: "11", ariaLabel: "Font size 11", icon: "11" },
                { id: "home.font.fontSize.12", label: "12", ariaLabel: "Font size 12", icon: "12" },
                { id: "home.font.fontSize.14", label: "14", ariaLabel: "Font size 14", icon: "14" },
                { id: "home.font.fontSize.16", label: "16", ariaLabel: "Font size 16", icon: "16" },
                { id: "home.font.fontSize.18", label: "18", ariaLabel: "Font size 18", icon: "18" },
                { id: "home.font.fontSize.20", label: "20", ariaLabel: "Font size 20", icon: "20" },
                { id: "home.font.fontSize.24", label: "24", ariaLabel: "Font size 24", icon: "24" },
                { id: "home.font.fontSize.28", label: "28", ariaLabel: "Font size 28", icon: "28" },
                { id: "home.font.fontSize.36", label: "36", ariaLabel: "Font size 36", icon: "36" },
                { id: "home.font.fontSize.48", label: "48", ariaLabel: "Font size 48", icon: "48" },
              ],
            },
            { id: "home.font.increaseFont", label: "Grow Font", ariaLabel: "Increase Font Size", icon: "A+", size: "icon" },
            { id: "home.font.decreaseFont", label: "Shrink Font", ariaLabel: "Decrease Font Size", icon: "A-", size: "icon" },
            { id: "home.font.bold", label: "Bold", ariaLabel: "Bold", icon: "B", kind: "toggle", size: "icon" },
            { id: "home.font.italic", label: "Italic", ariaLabel: "Italic", icon: "I", kind: "toggle", size: "icon" },
            { id: "home.font.underline", label: "Underline", ariaLabel: "Underline", icon: "U", kind: "toggle", size: "icon" },
            { id: "home.font.strikethrough", label: "Strike", ariaLabel: "Strikethrough", icon: "S̶", kind: "toggle", size: "icon" },
            { id: "home.font.subscript", label: "Subscript", ariaLabel: "Subscript", icon: "x₂", kind: "toggle", size: "icon" },
            { id: "home.font.superscript", label: "Superscript", ariaLabel: "Superscript", icon: "x²", kind: "toggle", size: "icon" },
            {
              id: "home.font.borders",
              label: "Borders",
              ariaLabel: "Borders",
              icon: "▦",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.font.borders.none", label: "No Border", ariaLabel: "No Border", icon: "▢" },
                { id: "home.font.borders.all", label: "All Borders", ariaLabel: "All Borders", icon: "▦" },
                { id: "home.font.borders.outside", label: "Outside Borders", ariaLabel: "Outside Borders", icon: "⬚" },
                { id: "home.font.borders.thickBox", label: "Thick Box Border", ariaLabel: "Thick Box Border", icon: "⬛" },
                { id: "home.font.borders.bottom", label: "Bottom Border", ariaLabel: "Bottom Border", icon: "▁" },
                { id: "home.font.borders.top", label: "Top Border", ariaLabel: "Top Border", icon: "▔" },
                { id: "home.font.borders.left", label: "Left Border", ariaLabel: "Left Border", icon: "▏" },
                { id: "home.font.borders.right", label: "Right Border", ariaLabel: "Right Border", icon: "▕" },
              ],
            },
            {
              id: "home.font.fillColor",
              label: "Fill",
              ariaLabel: "Fill Color",
              icon: "🪣",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.font.fillColor.noFill", label: "No Fill", ariaLabel: "No Fill", icon: "⬜" },
                { id: "home.font.fillColor.lightGray", label: "Light Gray", ariaLabel: "Light Gray Fill", icon: "⬚" },
                { id: "home.font.fillColor.yellow", label: "Yellow", ariaLabel: "Yellow Fill", icon: "🟨" },
                { id: "home.font.fillColor.green", label: "Green", ariaLabel: "Green Fill", icon: "🟩" },
                { id: "home.font.fillColor.red", label: "Red", ariaLabel: "Red Fill", icon: "🟥" },
              ],
            },
            {
              id: "home.font.fontColor",
              label: "Color",
              ariaLabel: "Font Color",
              icon: "🎨",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.font.fontColor.automatic", label: "Automatic", ariaLabel: "Automatic Font Color", icon: "A" },
                { id: "home.font.fontColor.black", label: "Black", ariaLabel: "Black Font Color", icon: "⬛" },
                { id: "home.font.fontColor.blue", label: "Blue", ariaLabel: "Blue Font Color", icon: "🟦" },
                { id: "home.font.fontColor.red", label: "Red", ariaLabel: "Red Font Color", icon: "🟥" },
                { id: "home.font.fontColor.green", label: "Green", ariaLabel: "Green Font Color", icon: "🟩" },
              ],
            },
            { id: "home.font.clearFormatting", label: "Clear", ariaLabel: "Clear Formatting", icon: "🧼", kind: "dropdown", size: "icon" },
          ],
        },
        {
          id: "home.alignment",
          label: "Alignment",
          buttons: [
            { id: "home.alignment.topAlign", label: "Top", ariaLabel: "Top Align", icon: "⬆", size: "icon" },
            { id: "home.alignment.middleAlign", label: "Middle", ariaLabel: "Middle Align", icon: "↕", size: "icon" },
            { id: "home.alignment.bottomAlign", label: "Bottom", ariaLabel: "Bottom Align", icon: "⬇", size: "icon" },
            { id: "home.alignment.alignLeft", label: "Left", ariaLabel: "Align Left", icon: "⬅", size: "icon" },
            { id: "home.alignment.center", label: "Center", ariaLabel: "Center", icon: "↔", size: "icon" },
            { id: "home.alignment.alignRight", label: "Right", ariaLabel: "Align Right", icon: "➡", size: "icon" },
            { id: "home.alignment.orientation", label: "Orientation", ariaLabel: "Orientation", icon: "↻", kind: "dropdown", size: "icon" },
            { id: "home.alignment.wrapText", label: "Wrap Text", ariaLabel: "Wrap Text", icon: "↩", kind: "toggle", size: "small" },
            {
              id: "home.alignment.mergeCenter",
              label: "Merge & Center",
              ariaLabel: "Merge and Center",
              icon: "⊞",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.alignment.mergeCenter.mergeCenter", label: "Merge & Center", ariaLabel: "Merge and Center", icon: "⊞" },
                { id: "home.alignment.mergeCenter.mergeAcross", label: "Merge Across", ariaLabel: "Merge Across", icon: "↔" },
                { id: "home.alignment.mergeCenter.mergeCells", label: "Merge Cells", ariaLabel: "Merge Cells", icon: "▦" },
                { id: "home.alignment.mergeCenter.unmergeCells", label: "Unmerge Cells", ariaLabel: "Unmerge Cells", icon: "✖" },
              ],
            },
            { id: "home.alignment.increaseIndent", label: "Indent", ariaLabel: "Increase Indent", icon: "⇥", size: "icon" },
            { id: "home.alignment.decreaseIndent", label: "Outdent", ariaLabel: "Decrease Indent", icon: "⇤", size: "icon" },
          ],
        },
        {
          id: "home.number",
          label: "Number",
          buttons: [
            {
              id: "home.number.numberFormat",
              label: "General",
              ariaLabel: "Number Format",
              icon: "123",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.number.numberFormat.general", label: "General", ariaLabel: "General", icon: "123" },
                { id: "home.number.numberFormat.number", label: "Number", ariaLabel: "Number", icon: "0.00" },
                { id: "home.number.numberFormat.currency", label: "Currency", ariaLabel: "Currency", icon: "$" },
                { id: "home.number.numberFormat.accounting", label: "Accounting", ariaLabel: "Accounting", icon: "$" },
                { id: "home.number.numberFormat.shortDate", label: "Short Date", ariaLabel: "Short Date", icon: "📅" },
                { id: "home.number.numberFormat.longDate", label: "Long Date", ariaLabel: "Long Date", icon: "📅" },
                { id: "home.number.numberFormat.time", label: "Time", ariaLabel: "Time", icon: "🕒" },
                { id: "home.number.numberFormat.percentage", label: "Percentage", ariaLabel: "Percentage", icon: "%" },
                { id: "home.number.numberFormat.fraction", label: "Fraction", ariaLabel: "Fraction", icon: "½" },
                { id: "home.number.numberFormat.scientific", label: "Scientific", ariaLabel: "Scientific", icon: "E" },
                { id: "home.number.numberFormat.text", label: "Text", ariaLabel: "Text", icon: "Aa" },
              ],
            },
            {
              id: "home.number.accounting",
              label: "Accounting",
              ariaLabel: "Accounting Number Format",
              icon: "$",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.number.accounting.usd", label: "$ (Dollar)", ariaLabel: "Dollar", icon: "$" },
                { id: "home.number.accounting.eur", label: "€ (Euro)", ariaLabel: "Euro", icon: "€" },
                { id: "home.number.accounting.gbp", label: "£ (Pound)", ariaLabel: "Pound", icon: "£" },
                { id: "home.number.accounting.jpy", label: "¥ (Yen)", ariaLabel: "Yen", icon: "¥" },
              ],
            },
            { id: "home.number.percent", label: "Percent", ariaLabel: "Percent Style", icon: "%", size: "icon" },
            { id: "home.number.date", label: "Date", ariaLabel: "Date", icon: "📅", size: "icon" },
            { id: "home.number.comma", label: "Comma", ariaLabel: "Comma Style", icon: ",", size: "icon" },
            { id: "home.number.increaseDecimal", label: "Inc Decimal", ariaLabel: "Increase Decimal", icon: ".0→", size: "icon" },
            { id: "home.number.decreaseDecimal", label: "Dec Decimal", ariaLabel: "Decrease Decimal", icon: "←.0", size: "icon" },
            {
              id: "home.number.moreFormats",
              label: "More",
              ariaLabel: "More Number Formats",
              icon: "⋯",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.number.moreFormats.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", icon: "⚙️" },
                { id: "home.number.moreFormats.custom", label: "Custom…", ariaLabel: "Custom Number Format", icon: "✎" },
              ],
            },
            { id: "home.number.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", icon: "⚙️", size: "small", testId: "ribbon-format-cells" },
          ],
        },
        {
          id: "home.styles",
          label: "Styles",
          buttons: [
            {
              id: "home.styles.conditionalFormatting",
              label: "Conditional Formatting",
              ariaLabel: "Conditional Formatting",
              icon: "📊",
              kind: "dropdown",
              size: "large",
            },
            { id: "home.styles.formatAsTable", label: "Format as Table", ariaLabel: "Format as Table", icon: "📋", kind: "dropdown", size: "large" },
            { id: "home.styles.cellStyles", label: "Cell Styles", ariaLabel: "Cell Styles", icon: "🎨", kind: "dropdown", size: "large" },
          ],
        },
        {
          id: "home.cells",
          label: "Cells",
          buttons: [
            {
              id: "home.cells.insert",
              label: "Insert",
              ariaLabel: "Insert Cells",
              icon: "⊕",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.insert.insertCells", label: "Insert Cells…", ariaLabel: "Insert Cells", icon: "⊞" },
                { id: "home.cells.insert.insertSheetRows", label: "Insert Sheet Rows", ariaLabel: "Insert Sheet Rows", icon: "↧" },
                { id: "home.cells.insert.insertSheetColumns", label: "Insert Sheet Columns", ariaLabel: "Insert Sheet Columns", icon: "↦" },
                { id: "home.cells.insert.insertSheet", label: "Insert Sheet", ariaLabel: "Insert Sheet", icon: "📄" },
              ],
            },
            {
              id: "home.cells.delete",
              label: "Delete",
              ariaLabel: "Delete Cells",
              icon: "⊖",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.delete.deleteCells", label: "Delete Cells…", ariaLabel: "Delete Cells", icon: "⊟" },
                { id: "home.cells.delete.deleteSheetRows", label: "Delete Sheet Rows", ariaLabel: "Delete Sheet Rows", icon: "↥" },
                { id: "home.cells.delete.deleteSheetColumns", label: "Delete Sheet Columns", ariaLabel: "Delete Sheet Columns", icon: "↤" },
                { id: "home.cells.delete.deleteSheet", label: "Delete Sheet", ariaLabel: "Delete Sheet", icon: "🗑" },
              ],
            },
            {
              id: "home.cells.format",
              label: "Format",
              ariaLabel: "Format Cells",
              icon: "⊡",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.format.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", icon: "⚙️" },
                { id: "home.cells.format.rowHeight", label: "Row Height…", ariaLabel: "Row Height", icon: "↕" },
                { id: "home.cells.format.columnWidth", label: "Column Width…", ariaLabel: "Column Width", icon: "↔" },
                { id: "home.cells.format.organizeSheets", label: "Organize Sheets", ariaLabel: "Organize Sheets", icon: "🗂" },
              ],
            },
          ],
        },
        {
          id: "home.editing",
          label: "Editing",
          buttons: [
            {
              id: "home.editing.autoSum",
              label: "AutoSum",
              ariaLabel: "AutoSum",
              icon: "Σ",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.autoSum.sum", label: "Sum", ariaLabel: "Sum", icon: "Σ" },
                { id: "home.editing.autoSum.average", label: "Average", ariaLabel: "Average", icon: "x̄" },
                { id: "home.editing.autoSum.countNumbers", label: "Count Numbers", ariaLabel: "Count Numbers", icon: "#" },
                { id: "home.editing.autoSum.max", label: "Max", ariaLabel: "Max", icon: "↑" },
                { id: "home.editing.autoSum.min", label: "Min", ariaLabel: "Min", icon: "↓" },
                { id: "home.editing.autoSum.moreFunctions", label: "More Functions…", ariaLabel: "More Functions", icon: "fx" },
              ],
            },
            {
              id: "home.editing.fill",
              label: "Fill",
              ariaLabel: "Fill",
              icon: "↓",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.fill.down", label: "Down", ariaLabel: "Fill Down", icon: "↓" },
                { id: "home.editing.fill.right", label: "Right", ariaLabel: "Fill Right", icon: "→" },
                { id: "home.editing.fill.up", label: "Up", ariaLabel: "Fill Up", icon: "↑" },
                { id: "home.editing.fill.left", label: "Left", ariaLabel: "Fill Left", icon: "←" },
                { id: "home.editing.fill.series", label: "Series…", ariaLabel: "Series", icon: "⋯" },
              ],
            },
            {
              id: "home.editing.clear",
              label: "Clear",
              ariaLabel: "Clear",
              icon: "⌫",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.clear.clearAll", label: "Clear All", ariaLabel: "Clear All", icon: "🧹" },
                { id: "home.editing.clear.clearFormats", label: "Clear Formats", ariaLabel: "Clear Formats", icon: "🎨" },
                { id: "home.editing.clear.clearContents", label: "Clear Contents", ariaLabel: "Clear Contents", icon: "⌫" },
                { id: "home.editing.clear.clearComments", label: "Clear Comments", ariaLabel: "Clear Comments", icon: "💬" },
                { id: "home.editing.clear.clearHyperlinks", label: "Clear Hyperlinks", ariaLabel: "Clear Hyperlinks", icon: "🔗" },
              ],
            },
            {
              id: "home.editing.sortFilter",
              label: "Sort & Filter",
              ariaLabel: "Sort and Filter",
              icon: "⇅",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", icon: "A→Z" },
                { id: "home.editing.sortFilter.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", icon: "Z→A" },
                { id: "home.editing.sortFilter.customSort", label: "Custom Sort…", ariaLabel: "Custom Sort", icon: "⚙️" },
                { id: "home.editing.sortFilter.filter", label: "Filter", ariaLabel: "Filter", icon: "⏷" },
                { id: "home.editing.sortFilter.clear", label: "Clear", ariaLabel: "Clear", icon: "✖" },
                { id: "home.editing.sortFilter.reapply", label: "Reapply", ariaLabel: "Reapply", icon: "⟳" },
              ],
            },
            {
              id: "home.editing.findSelect",
              label: "Find & Select",
              ariaLabel: "Find and Select",
              icon: "⌕",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.findSelect.find", label: "Find", ariaLabel: "Find", icon: "⌕", testId: "ribbon-find" },
                { id: "home.editing.findSelect.replace", label: "Replace", ariaLabel: "Replace", icon: "⎘", testId: "ribbon-replace" },
                { id: "home.editing.findSelect.goTo", label: "Go To", ariaLabel: "Go To", icon: "↗", testId: "ribbon-goto" },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "insert",
      label: "Insert",
      groups: [
        {
          id: "insert.tables",
          label: "Tables",
          buttons: [
            { id: "insert.tables.pivotTable", label: "PivotTable", ariaLabel: "PivotTable", icon: "📊", kind: "dropdown", size: "large" },
            { id: "insert.tables.recommendedPivotTables", label: "Recommended PivotTables", ariaLabel: "Recommended PivotTables", icon: "✨", kind: "dropdown" },
            { id: "insert.tables.table", label: "Table", ariaLabel: "Table", icon: "▦", size: "large" },
          ],
        },
        {
          id: "insert.pivotcharts",
          label: "PivotCharts",
          buttons: [
            { id: "insert.pivotcharts.pivotChart", label: "PivotChart", ariaLabel: "PivotChart", icon: "📈", kind: "dropdown", size: "large" },
            { id: "insert.pivotcharts.recommendedPivotCharts", label: "Recommended PivotCharts", ariaLabel: "Recommended PivotCharts", icon: "✨", kind: "dropdown" },
          ],
        },
        {
          id: "insert.illustrations",
          label: "Illustrations",
          buttons: [
            { id: "insert.illustrations.pictures", label: "Pictures", ariaLabel: "Pictures", icon: "🖼", kind: "dropdown" },
            { id: "insert.illustrations.onlinePictures", label: "Online Pictures", ariaLabel: "Online Pictures", icon: "🌐", kind: "dropdown" },
            { id: "insert.illustrations.shapes", label: "Shapes", ariaLabel: "Shapes", icon: "⬛", kind: "dropdown" },
            { id: "insert.illustrations.icons", label: "Icons", ariaLabel: "Icons", icon: "⭐", kind: "dropdown" },
            { id: "insert.illustrations.smartArt", label: "SmartArt", ariaLabel: "SmartArt", icon: "🧩", kind: "dropdown" },
            { id: "insert.illustrations.screenshot", label: "Screenshot", ariaLabel: "Screenshot", icon: "📸", kind: "dropdown" },
          ],
        },
        {
          id: "insert.addins",
          label: "Add-ins",
          buttons: [
            { id: "insert.addins.getAddins", label: "Get Add-ins", ariaLabel: "Get Add-ins", icon: "➕", kind: "dropdown" },
            { id: "insert.addins.myAddins", label: "My Add-ins", ariaLabel: "My Add-ins", icon: "🧩", kind: "dropdown" },
          ],
        },
        {
          id: "insert.charts",
          label: "Charts",
          buttons: [
            { id: "insert.charts.recommendedCharts", label: "Recommended Charts", ariaLabel: "Recommended Charts", icon: "✨", kind: "dropdown", size: "large" },
            { id: "insert.charts.column", label: "Column", ariaLabel: "Insert Column or Bar Chart", icon: "▮▮", kind: "dropdown" },
            { id: "insert.charts.line", label: "Line", ariaLabel: "Insert Line or Area Chart", icon: "📈", kind: "dropdown" },
            { id: "insert.charts.pie", label: "Pie", ariaLabel: "Insert Pie or Doughnut Chart", icon: "◔", kind: "dropdown" },
            { id: "insert.charts.bar", label: "Bar", ariaLabel: "Insert Bar Chart", icon: "▭", kind: "dropdown" },
            { id: "insert.charts.area", label: "Area", ariaLabel: "Insert Area Chart", icon: "⛰", kind: "dropdown" },
            { id: "insert.charts.scatter", label: "Scatter", ariaLabel: "Insert Scatter (X, Y) Chart", icon: "⋯", kind: "dropdown" },
            { id: "insert.charts.map", label: "Map", ariaLabel: "Insert Map Chart", icon: "🗺", kind: "dropdown" },
            { id: "insert.charts.histogram", label: "Histogram", ariaLabel: "Insert Statistic Chart (Histogram, Pareto)", icon: "▁▃▆", kind: "dropdown" },
            { id: "insert.charts.waterfall", label: "Waterfall", ariaLabel: "Insert Waterfall Chart", icon: "💧", kind: "dropdown" },
            { id: "insert.charts.treemap", label: "Treemap", ariaLabel: "Insert Hierarchy Chart (Treemap)", icon: "🧩", kind: "dropdown" },
            { id: "insert.charts.sunburst", label: "Sunburst", ariaLabel: "Insert Hierarchy Chart (Sunburst)", icon: "☀️", kind: "dropdown" },
            { id: "insert.charts.funnel", label: "Funnel", ariaLabel: "Insert Funnel Chart", icon: "⏬", kind: "dropdown" },
            { id: "insert.charts.boxWhisker", label: "Box & Whisker", ariaLabel: "Insert Box and Whisker Chart", icon: "▣", kind: "dropdown" },
            { id: "insert.charts.radar", label: "Radar", ariaLabel: "Insert Radar Chart", icon: "🕸", kind: "dropdown" },
            { id: "insert.charts.surface", label: "Surface", ariaLabel: "Insert Surface Chart", icon: "🗻", kind: "dropdown" },
            { id: "insert.charts.stock", label: "Stock", ariaLabel: "Insert Stock Chart", icon: "💹", kind: "dropdown" },
            { id: "insert.charts.combo", label: "Combo", ariaLabel: "Insert Combo Chart", icon: "🔀", kind: "dropdown" },
            { id: "insert.charts.pivotChart", label: "PivotChart", ariaLabel: "PivotChart", icon: "📊", kind: "dropdown" },
          ],
        },
        {
          id: "insert.tours",
          label: "Tours",
          buttons: [
            { id: "insert.tours.3dMap", label: "3D Map", ariaLabel: "3D Map", icon: "🌍", kind: "dropdown", size: "large" },
            { id: "insert.tours.launchTour", label: "Launch Tour", ariaLabel: "Launch Tour", icon: "🚀", kind: "dropdown" },
          ],
        },
        {
          id: "insert.sparklines",
          label: "Sparklines",
          buttons: [
            { id: "insert.sparklines.line", label: "Line", ariaLabel: "Insert Line Sparkline", icon: "╱", kind: "dropdown" },
            { id: "insert.sparklines.column", label: "Column", ariaLabel: "Insert Column Sparkline", icon: "▮", kind: "dropdown" },
            { id: "insert.sparklines.winLoss", label: "Win/Loss", ariaLabel: "Insert Win/Loss Sparkline", icon: "±", kind: "dropdown" },
          ],
        },
        {
          id: "insert.filters",
          label: "Filters",
          buttons: [
            { id: "insert.filters.slicer", label: "Slicer", ariaLabel: "Insert Slicer", icon: "🔪", kind: "dropdown" },
            { id: "insert.filters.timeline", label: "Timeline", ariaLabel: "Insert Timeline", icon: "🕒", kind: "dropdown" },
          ],
        },
        {
          id: "insert.links",
          label: "Links",
          buttons: [{ id: "insert.links.link", label: "Link", ariaLabel: "Insert Link", icon: "🔗", kind: "dropdown", size: "large" }],
        },
        {
          id: "insert.comments",
          label: "Comments",
          buttons: [
            { id: "insert.comments.comment", label: "Comment", ariaLabel: "Insert Comment", icon: "💬", kind: "dropdown", size: "large" },
            { id: "insert.comments.note", label: "Note", ariaLabel: "Insert Note", icon: "🗒", kind: "dropdown" },
          ],
        },
        {
          id: "insert.text",
          label: "Text",
          buttons: [
            { id: "insert.text.textBox", label: "Text Box", ariaLabel: "Insert Text Box", icon: "📝", kind: "dropdown" },
            { id: "insert.text.headerFooter", label: "Header & Footer", ariaLabel: "Header and Footer", icon: "📄", kind: "dropdown" },
            { id: "insert.text.wordArt", label: "WordArt", ariaLabel: "WordArt", icon: "𝒜", kind: "dropdown" },
            { id: "insert.text.signatureLine", label: "Signature Line", ariaLabel: "Signature Line", icon: "✍️", kind: "dropdown" },
            { id: "insert.text.object", label: "Object", ariaLabel: "Object", icon: "🧱", kind: "dropdown" },
          ],
        },
        {
          id: "insert.equations",
          label: "Equations",
          buttons: [
            { id: "insert.equations.equation", label: "Equation", ariaLabel: "Insert Equation", icon: "∑", kind: "dropdown", size: "large" },
            { id: "insert.equations.inkEquation", label: "Ink Equation", ariaLabel: "Ink Equation", icon: "✒️", kind: "dropdown" },
          ],
        },
        {
          id: "insert.symbols",
          label: "Symbols",
          buttons: [
            { id: "insert.symbols.equation", label: "Equation", ariaLabel: "Insert Equation", icon: "∑", kind: "dropdown" },
            { id: "insert.symbols.symbol", label: "Symbol", ariaLabel: "Insert Symbol", icon: "Ω", kind: "dropdown" },
          ],
        },
      ],
    },
    {
      id: "pageLayout",
      label: "Page Layout",
      groups: [
        {
          id: "pageLayout.themes",
          label: "Themes",
          buttons: [
            { id: "pageLayout.themes.themes", label: "Themes", ariaLabel: "Themes", icon: "🎛", kind: "dropdown", size: "large" },
            { id: "pageLayout.themes.colors", label: "Colors", ariaLabel: "Colors", icon: "🎨", kind: "dropdown" },
            { id: "pageLayout.themes.fonts", label: "Fonts", ariaLabel: "Fonts", icon: "🔤", kind: "dropdown" },
            { id: "pageLayout.themes.effects", label: "Effects", ariaLabel: "Effects", icon: "✨", kind: "dropdown" },
          ],
        },
        {
          id: "pageLayout.pageSetup",
          label: "Page Setup",
          buttons: [
            { id: "pageLayout.pageSetup.margins", label: "Margins", ariaLabel: "Margins", icon: "📏", kind: "dropdown" },
            { id: "pageLayout.pageSetup.orientation", label: "Orientation", ariaLabel: "Orientation", icon: "↔", kind: "dropdown" },
            { id: "pageLayout.pageSetup.size", label: "Size", ariaLabel: "Size", icon: "📄", kind: "dropdown" },
            { id: "pageLayout.pageSetup.printArea", label: "Print Area", ariaLabel: "Print Area", icon: "🖨", kind: "dropdown" },
            { id: "pageLayout.pageSetup.breaks", label: "Breaks", ariaLabel: "Breaks", icon: "⤶", kind: "dropdown" },
            { id: "pageLayout.pageSetup.background", label: "Background", ariaLabel: "Background", icon: "🖼", kind: "dropdown" },
            { id: "pageLayout.pageSetup.printTitles", label: "Print Titles", ariaLabel: "Print Titles", icon: "🏷", kind: "dropdown" },
          ],
        },
        {
          id: "pageLayout.scaleToFit",
          label: "Scale to Fit",
          buttons: [
            { id: "pageLayout.scaleToFit.width", label: "Width", ariaLabel: "Width", icon: "↔", kind: "dropdown" },
            { id: "pageLayout.scaleToFit.height", label: "Height", ariaLabel: "Height", icon: "↕", kind: "dropdown" },
            { id: "pageLayout.scaleToFit.scale", label: "Scale", ariaLabel: "Scale", icon: "🔍", kind: "dropdown" },
          ],
        },
        {
          id: "pageLayout.sheetOptions",
          label: "Sheet Options",
          buttons: [
            { id: "pageLayout.sheetOptions.gridlinesView", label: "Gridlines View", ariaLabel: "View Gridlines", icon: "▦", kind: "toggle", size: "small", defaultPressed: true },
            { id: "pageLayout.sheetOptions.gridlinesPrint", label: "Gridlines Print", ariaLabel: "Print Gridlines", icon: "🖨", kind: "toggle", size: "small", defaultPressed: false },
            { id: "pageLayout.sheetOptions.headingsView", label: "Headings View", ariaLabel: "View Headings", icon: "A1", kind: "toggle", size: "small", defaultPressed: true },
            { id: "pageLayout.sheetOptions.headingsPrint", label: "Headings Print", ariaLabel: "Print Headings", icon: "🖨", kind: "toggle", size: "small", defaultPressed: false },
          ],
        },
        {
          id: "pageLayout.arrange",
          label: "Arrange",
          buttons: [
            { id: "pageLayout.arrange.bringForward", label: "Bring Forward", ariaLabel: "Bring Forward", icon: "⬆", kind: "dropdown" },
            { id: "pageLayout.arrange.sendBackward", label: "Send Backward", ariaLabel: "Send Backward", icon: "⬇", kind: "dropdown" },
            { id: "pageLayout.arrange.selectionPane", label: "Selection Pane", ariaLabel: "Selection Pane", icon: "📋" },
            { id: "pageLayout.arrange.align", label: "Align", ariaLabel: "Align", icon: "📐", kind: "dropdown" },
            { id: "pageLayout.arrange.group", label: "Group", ariaLabel: "Group", icon: "🔗", kind: "dropdown" },
            { id: "pageLayout.arrange.rotate", label: "Rotate", ariaLabel: "Rotate", icon: "↻", kind: "dropdown" },
          ],
        },
      ],
    },
    {
      id: "formulas",
      label: "Formulas",
      groups: [
        {
          id: "formulas.functionLibrary",
          label: "Function Library",
          buttons: [
            { id: "formulas.functionLibrary.insertFunction", label: "Insert Function", ariaLabel: "Insert Function", icon: "fx", kind: "dropdown", size: "large" },
            { id: "formulas.functionLibrary.autoSum", label: "AutoSum", ariaLabel: "AutoSum", icon: "Σ", kind: "dropdown" },
            { id: "formulas.functionLibrary.recentlyUsed", label: "Recently Used", ariaLabel: "Recently Used", icon: "🕘", kind: "dropdown" },
            { id: "formulas.functionLibrary.financial", label: "Financial", ariaLabel: "Financial", icon: "$", kind: "dropdown" },
            { id: "formulas.functionLibrary.logical", label: "Logical", ariaLabel: "Logical", icon: "∧", kind: "dropdown" },
            { id: "formulas.functionLibrary.text", label: "Text", ariaLabel: "Text", icon: "Aa", kind: "dropdown" },
            { id: "formulas.functionLibrary.dateTime", label: "Date & Time", ariaLabel: "Date and Time", icon: "📅", kind: "dropdown" },
            { id: "formulas.functionLibrary.lookupReference", label: "Lookup & Reference", ariaLabel: "Lookup and Reference", icon: "🔎", kind: "dropdown" },
            { id: "formulas.functionLibrary.mathTrig", label: "Math & Trig", ariaLabel: "Math and Trig", icon: "π", kind: "dropdown" },
            { id: "formulas.functionLibrary.moreFunctions", label: "More Functions", ariaLabel: "More Functions", icon: "➕", kind: "dropdown" },
          ],
        },
        {
          id: "formulas.definedNames",
          label: "Defined Names",
          buttons: [
            { id: "formulas.definedNames.nameManager", label: "Name Manager", ariaLabel: "Name Manager", icon: "🏷", kind: "dropdown", size: "large" },
            { id: "formulas.definedNames.defineName", label: "Define Name", ariaLabel: "Define Name", icon: "➕", kind: "dropdown" },
            { id: "formulas.definedNames.useInFormula", label: "Use in Formula", ariaLabel: "Use in Formula", icon: "fx", kind: "dropdown" },
            { id: "formulas.definedNames.createFromSelection", label: "Create from Selection", ariaLabel: "Create from Selection", icon: "▦", kind: "dropdown" },
          ],
        },
        {
          id: "formulas.formulaAuditing",
          label: "Formula Auditing",
          buttons: [
            { id: "formulas.formulaAuditing.tracePrecedents", label: "Trace Precedents", ariaLabel: "Trace Precedents", icon: "⬅", size: "small" },
            { id: "formulas.formulaAuditing.traceDependents", label: "Trace Dependents", ariaLabel: "Trace Dependents", icon: "➡", size: "small" },
            { id: "formulas.formulaAuditing.removeArrows", label: "Remove Arrows", ariaLabel: "Remove Arrows", icon: "✖", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.showFormulas", label: "Show Formulas", ariaLabel: "Show Formulas", icon: "ƒx", kind: "toggle", size: "small" },
            { id: "formulas.formulaAuditing.errorChecking", label: "Error Checking", ariaLabel: "Error Checking", icon: "⚠", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.evaluateFormula", label: "Evaluate Formula", ariaLabel: "Evaluate Formula", icon: "🧮", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.watchWindow", label: "Watch Window", ariaLabel: "Watch Window", icon: "👁", kind: "dropdown", size: "small" },
          ],
        },
        {
          id: "formulas.calculation",
          label: "Calculation",
          buttons: [
            { id: "formulas.calculation.calculationOptions", label: "Calculation Options", ariaLabel: "Calculation Options", icon: "⚙️", kind: "dropdown", size: "large" },
            { id: "formulas.calculation.calculateNow", label: "Calculate Now", ariaLabel: "Calculate Now", icon: "⟳", size: "small" },
            { id: "formulas.calculation.calculateSheet", label: "Calculate Sheet", ariaLabel: "Calculate Sheet", icon: "⟲", size: "small" },
          ],
        },
        {
          id: "formulas.solutions",
          label: "Solutions",
          buttons: [
            { id: "formulas.solutions.solver", label: "Solver", ariaLabel: "Solver", icon: "🧩", kind: "dropdown", size: "large" },
            { id: "formulas.solutions.analysisToolPak", label: "Analysis ToolPak", ariaLabel: "Analysis ToolPak", icon: "🧰", kind: "dropdown" },
          ],
        },
      ],
    },
    {
      id: "data",
      label: "Data",
      groups: [
        {
          id: "data.getTransform",
          label: "Get & Transform Data",
          buttons: [
            { id: "data.getTransform.getData", label: "Get Data", ariaLabel: "Get Data", icon: "⬇", kind: "dropdown", size: "large" },
            { id: "data.getTransform.recentSources", label: "Recent Sources", ariaLabel: "Recent Sources", icon: "🕘", kind: "dropdown" },
            { id: "data.getTransform.existingConnections", label: "Existing Connections", ariaLabel: "Existing Connections", icon: "🔗", kind: "dropdown" },
          ],
        },
        {
          id: "data.queriesConnections",
          label: "Queries & Connections",
          buttons: [
            { id: "data.queriesConnections.refreshAll", label: "Refresh All", ariaLabel: "Refresh All", icon: "⟳", kind: "dropdown", size: "large" },
            { id: "data.queriesConnections.queriesConnections", label: "Queries & Connections", ariaLabel: "Queries and Connections", icon: "🗂", kind: "toggle", defaultPressed: false },
            { id: "data.queriesConnections.properties", label: "Properties", ariaLabel: "Properties", icon: "⚙️", kind: "dropdown" },
          ],
        },
        {
          id: "data.sortFilter",
          label: "Sort & Filter",
          buttons: [
            { id: "data.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", icon: "A→Z" },
            { id: "data.sortFilter.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", icon: "Z→A" },
            { id: "data.sortFilter.sort", label: "Sort", ariaLabel: "Sort", icon: "⇅", kind: "dropdown" },
            { id: "data.sortFilter.filter", label: "Filter", ariaLabel: "Filter", icon: "⏷", kind: "toggle" },
            { id: "data.sortFilter.clear", label: "Clear", ariaLabel: "Clear", icon: "✖" },
            { id: "data.sortFilter.reapply", label: "Reapply", ariaLabel: "Reapply", icon: "⟳" },
            { id: "data.sortFilter.advanced", label: "Advanced", ariaLabel: "Advanced", icon: "⚙️", kind: "dropdown" },
          ],
        },
        {
          id: "data.dataTools",
          label: "Data Tools",
          buttons: [
            { id: "data.dataTools.textToColumns", label: "Text to Columns", ariaLabel: "Text to Columns", icon: "⇥", kind: "dropdown" },
            { id: "data.dataTools.flashFill", label: "Flash Fill", ariaLabel: "Flash Fill", icon: "⚡" },
            { id: "data.dataTools.removeDuplicates", label: "Remove Duplicates", ariaLabel: "Remove Duplicates", icon: "🗑", kind: "dropdown" },
            { id: "data.dataTools.dataValidation", label: "Data Validation", ariaLabel: "Data Validation", icon: "✅", kind: "dropdown" },
            { id: "data.dataTools.consolidate", label: "Consolidate", ariaLabel: "Consolidate", icon: "🧩", kind: "dropdown" },
            { id: "data.dataTools.relationships", label: "Relationships", ariaLabel: "Relationships", icon: "🔗", kind: "dropdown" },
            { id: "data.dataTools.manageDataModel", label: "Manage Data Model", ariaLabel: "Manage Data Model", icon: "🧠", kind: "dropdown" },
          ],
        },
        {
          id: "data.forecast",
          label: "Forecast",
          buttons: [
            { id: "data.forecast.whatIfAnalysis", label: "What-If Analysis", ariaLabel: "What-If Analysis", icon: "❓", kind: "dropdown", size: "large" },
            { id: "data.forecast.forecastSheet", label: "Forecast Sheet", ariaLabel: "Forecast Sheet", icon: "📈", kind: "dropdown" },
          ],
        },
        {
          id: "data.outline",
          label: "Outline",
          buttons: [
            { id: "data.outline.group", label: "Group", ariaLabel: "Group", icon: "➕", kind: "dropdown" },
            { id: "data.outline.ungroup", label: "Ungroup", ariaLabel: "Ungroup", icon: "➖", kind: "dropdown" },
            { id: "data.outline.subtotal", label: "Subtotal", ariaLabel: "Subtotal", icon: "Σ", kind: "dropdown" },
            { id: "data.outline.showDetail", label: "Show Detail", ariaLabel: "Show Detail", icon: "＋" },
            { id: "data.outline.hideDetail", label: "Hide Detail", ariaLabel: "Hide Detail", icon: "−" },
          ],
        },
        {
          id: "data.dataTypes",
          label: "Data Types",
          buttons: [
            { id: "data.dataTypes.stocks", label: "Stocks", ariaLabel: "Stocks", icon: "📈", kind: "dropdown", size: "large" },
            { id: "data.dataTypes.geography", label: "Geography", ariaLabel: "Geography", icon: "🌎", kind: "dropdown", size: "large" },
          ],
        },
      ],
    },
    {
      id: "review",
      label: "Review",
      groups: [
        {
          id: "review.proofing",
          label: "Proofing",
          buttons: [
            { id: "review.proofing.spelling", label: "Spelling", ariaLabel: "Spelling", icon: "✔", kind: "dropdown", size: "large" },
            { id: "review.proofing.accessibility", label: "Check Accessibility", ariaLabel: "Check Accessibility", icon: "♿", kind: "dropdown" },
            { id: "review.proofing.smartLookup", label: "Smart Lookup", ariaLabel: "Smart Lookup", icon: "🔎", kind: "dropdown" },
          ],
        },
        {
          id: "review.comments",
          label: "Comments",
          buttons: [
            { id: "review.comments.newComment", label: "New Comment", ariaLabel: "New Comment", icon: "💬", size: "large" },
            { id: "review.comments.deleteComment", label: "Delete", ariaLabel: "Delete Comment", icon: "🗑", kind: "dropdown" },
            { id: "review.comments.previous", label: "Previous", ariaLabel: "Previous Comment", icon: "⬆" },
            { id: "review.comments.next", label: "Next", ariaLabel: "Next Comment", icon: "⬇" },
            { id: "review.comments.showComments", label: "Show Comments", ariaLabel: "Show Comments", icon: "👁", kind: "toggle" },
          ],
        },
        {
          id: "review.notes",
          label: "Notes",
          buttons: [
            { id: "review.notes.newNote", label: "New Note", ariaLabel: "New Note", icon: "🗒", kind: "dropdown", size: "large" },
            { id: "review.notes.showAllNotes", label: "Show All Notes", ariaLabel: "Show All Notes", icon: "👁", kind: "toggle" },
            { id: "review.notes.showHideNote", label: "Show/Hide Note", ariaLabel: "Show or Hide Note", icon: "🙈", kind: "toggle" },
          ],
        },
        {
          id: "review.protect",
          label: "Protect",
          buttons: [
            { id: "review.protect.protectSheet", label: "Protect Sheet", ariaLabel: "Protect Sheet", icon: "🔒", kind: "dropdown", size: "large" },
            { id: "review.protect.protectWorkbook", label: "Protect Workbook", ariaLabel: "Protect Workbook", icon: "🧰", kind: "dropdown" },
            { id: "review.protect.allowEditRanges", label: "Allow Edit Ranges", ariaLabel: "Allow Edit Ranges", icon: "✅", kind: "dropdown" },
          ],
        },
        {
          id: "review.ink",
          label: "Ink",
          buttons: [
            { id: "review.ink.startInking", label: "Start Inking", ariaLabel: "Start Inking", icon: "✒️", kind: "toggle", size: "large" },
          ],
        },
        {
          id: "review.language",
          label: "Language",
          buttons: [
            { id: "review.language.translate", label: "Translate", ariaLabel: "Translate", icon: "🌐", kind: "dropdown" },
            { id: "review.language.language", label: "Language", ariaLabel: "Language", icon: "🈯", kind: "dropdown" },
          ],
        },
        {
          id: "review.changes",
          label: "Changes",
          buttons: [
            { id: "review.changes.trackChanges", label: "Track Changes", ariaLabel: "Track Changes", icon: "📝", kind: "dropdown", size: "large" },
            { id: "review.changes.shareWorkbook", label: "Share Workbook", ariaLabel: "Share Workbook", icon: "👥", kind: "dropdown" },
            { id: "review.changes.protectShareWorkbook", label: "Protect and Share Workbook", ariaLabel: "Protect and Share Workbook", icon: "🔒", kind: "dropdown" },
          ],
        },
      ],
    },
    {
      id: "view",
      label: "View",
      groups: [
        {
          id: "view.workbookViews",
          label: "Workbook Views",
          buttons: [
            { id: "view.workbookViews.normal", label: "Normal", ariaLabel: "Normal View", icon: "▦", kind: "toggle", defaultPressed: true, size: "large" },
            { id: "view.workbookViews.pageBreakPreview", label: "Page Break Preview", ariaLabel: "Page Break Preview", icon: "⤶", kind: "toggle", size: "large" },
            { id: "view.workbookViews.pageLayout", label: "Page Layout", ariaLabel: "Page Layout View", icon: "📄", kind: "toggle", size: "large" },
            { id: "view.workbookViews.customViews", label: "Custom Views", ariaLabel: "Custom Views", icon: "👁", kind: "dropdown" },
          ],
        },
        {
          id: "view.show",
          label: "Show",
          buttons: [
            { id: "view.show.ruler", label: "Ruler", ariaLabel: "Ruler", icon: "📏", kind: "toggle", defaultPressed: false },
            { id: "view.show.gridlines", label: "Gridlines", ariaLabel: "Gridlines", icon: "▦", kind: "toggle", defaultPressed: true },
            { id: "view.show.formulaBar", label: "Formula Bar", ariaLabel: "Formula Bar", icon: "fx", kind: "toggle", defaultPressed: true },
            { id: "view.show.headings", label: "Headings", ariaLabel: "Headings", icon: "A1", kind: "toggle", defaultPressed: true },
          ],
        },
        {
          id: "view.zoom",
          label: "Zoom",
          buttons: [
            { id: "view.zoom.zoom", label: "Zoom", ariaLabel: "Zoom", icon: "🔍", kind: "dropdown", size: "large" },
            { id: "view.zoom.zoom100", label: "100%", ariaLabel: "Zoom to 100%", icon: "100%" },
            { id: "view.zoom.zoomToSelection", label: "Zoom to Selection", ariaLabel: "Zoom to Selection", icon: "🎯" },
          ],
        },
        {
          id: "view.window",
          label: "Window",
          buttons: [
            { id: "view.window.newWindow", label: "New Window", ariaLabel: "New Window", icon: "🪟", kind: "dropdown", size: "large" },
            { id: "view.window.arrangeAll", label: "Arrange All", ariaLabel: "Arrange All", icon: "🗔", kind: "dropdown" },
            { id: "view.window.freezePanes", label: "Freeze Panes", ariaLabel: "Freeze Panes", icon: "❄️", kind: "dropdown" },
            { id: "view.window.split", label: "Split", ariaLabel: "Split", icon: "➗", kind: "toggle" },
            { id: "view.window.hide", label: "Hide", ariaLabel: "Hide", icon: "🙈" },
            { id: "view.window.unhide", label: "Unhide", ariaLabel: "Unhide", icon: "👁" },
            { id: "view.window.viewSideBySide", label: "View Side by Side", ariaLabel: "View Side by Side", icon: "⧉", kind: "toggle" },
            { id: "view.window.synchronousScrolling", label: "Synchronous Scrolling", ariaLabel: "Synchronous Scrolling", icon: "⇵", kind: "toggle" },
            { id: "view.window.resetWindowPosition", label: "Reset Window Position", ariaLabel: "Reset Window Position", icon: "↺" },
            { id: "view.window.switchWindows", label: "Switch Windows", ariaLabel: "Switch Windows", icon: "🔁", kind: "dropdown" },
          ],
        },
        {
          id: "view.macros",
          label: "Macros",
          buttons: [
            { id: "view.macros.viewMacros", label: "View Macros", ariaLabel: "View Macros", icon: "📜", kind: "dropdown", size: "large" },
            { id: "view.macros.recordMacro", label: "Record Macro", ariaLabel: "Record Macro", icon: "⏺", kind: "dropdown" },
            { id: "view.macros.useRelativeReferences", label: "Use Relative References", ariaLabel: "Use Relative References", icon: "📍", kind: "toggle" },
          ],
        },
      ],
    },
    {
      id: "developer",
      label: "Developer",
      groups: [
        {
          id: "developer.code",
          label: "Code",
          buttons: [
            { id: "developer.code.visualBasic", label: "Visual Basic", ariaLabel: "Visual Basic", icon: "VB", size: "large" },
            { id: "developer.code.macros", label: "Macros", ariaLabel: "Macros", icon: "📜", kind: "dropdown", size: "large" },
            { id: "developer.code.recordMacro", label: "Record Macro", ariaLabel: "Record Macro", icon: "⏺", kind: "dropdown" },
            { id: "developer.code.useRelativeReferences", label: "Use Relative References", ariaLabel: "Use Relative References", icon: "📍", kind: "toggle" },
            { id: "developer.code.macroSecurity", label: "Macro Security", ariaLabel: "Macro Security", icon: "🔒", kind: "dropdown" },
          ],
        },
        {
          id: "developer.addins",
          label: "Add-ins",
          buttons: [
            { id: "developer.addins.addins", label: "Add-ins", ariaLabel: "Add-ins", icon: "🧩", kind: "dropdown", size: "large" },
            { id: "developer.addins.comAddins", label: "COM Add-ins", ariaLabel: "COM Add-ins", icon: "🔌", kind: "dropdown" },
          ],
        },
        {
          id: "developer.controls",
          label: "Controls",
          buttons: [
            { id: "developer.controls.insert", label: "Insert", ariaLabel: "Insert Control", icon: "➕", kind: "dropdown", size: "large" },
            { id: "developer.controls.designMode", label: "Design Mode", ariaLabel: "Design Mode", icon: "🎛", kind: "toggle" },
            { id: "developer.controls.properties", label: "Properties", ariaLabel: "Properties", icon: "⚙️", kind: "dropdown" },
            { id: "developer.controls.viewCode", label: "View Code", ariaLabel: "View Code", icon: "</>" },
            { id: "developer.controls.runDialog", label: "Run Dialog", ariaLabel: "Run Dialog", icon: "▶" },
          ],
        },
        {
          id: "developer.xml",
          label: "XML",
          buttons: [
            { id: "developer.xml.source", label: "Source", ariaLabel: "XML Source", icon: "XML", kind: "dropdown", size: "large" },
            { id: "developer.xml.mapProperties", label: "Map Properties", ariaLabel: "Map Properties", icon: "🗺", kind: "dropdown" },
            { id: "developer.xml.import", label: "Import", ariaLabel: "Import XML", icon: "⬇" },
            { id: "developer.xml.export", label: "Export", ariaLabel: "Export XML", icon: "⬆" },
            { id: "developer.xml.refreshData", label: "Refresh Data", ariaLabel: "Refresh Data", icon: "⟳" },
          ],
        },
      ],
    },
    {
      id: "help",
      label: "Help",
      groups: [
        {
          id: "help.support",
          label: "Support",
          buttons: [
            { id: "help.support.help", label: "Help", ariaLabel: "Help", icon: "❓", kind: "dropdown", size: "large" },
            { id: "help.support.training", label: "Training", ariaLabel: "Training", icon: "🎓", kind: "dropdown" },
            { id: "help.support.contactSupport", label: "Contact Support", ariaLabel: "Contact Support", icon: "☎️", kind: "dropdown" },
            { id: "help.support.feedback", label: "Feedback", ariaLabel: "Feedback", icon: "📝", kind: "dropdown" },
          ],
        },
      ],
    },
  ],
};
