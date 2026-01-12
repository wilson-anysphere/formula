import type { RibbonIconId } from "./icons/index.js";

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
   * Stable icon identifier.
   */
  iconId?: RibbonIconId;
  /**
   * Legacy text glyph fallback.
   *
   * Prefer `iconId` for consistent SVG icons.
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
   * Stable icon identifier.
   */
  iconId?: RibbonIconId;
  /**
   * Legacy text glyph fallback.
   *
   * Prefer `iconId` for consistent SVG icons.
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

export interface RibbonFileActions {
  newWorkbook?: () => void;
  openWorkbook?: () => void;
  saveWorkbook?: () => void;
  saveWorkbookAs?: () => void;
  versionHistory?: () => void;
  branchManager?: () => void;
  print?: () => void;
  pageSetup?: () => void;
  closeWindow?: () => void;
  quit?: () => void;
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
  /**
   * Optional File tab / backstage actions.
   *
   * The File tab is treated specially (Excel-style "backstage" view) and is
   * wired directly to app-level file operations in `apps/desktop/src/main.ts`.
   */
  fileActions?: RibbonFileActions;
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
            {
              id: "file.new.new",
              label: "New",
              ariaLabel: "New",
              iconId: "file",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "file.new.blankWorkbook", label: "Blank workbook", ariaLabel: "Blank workbook", iconId: "file" },
                { id: "file.new.templates", label: "Templates", ariaLabel: "Templates", iconId: "file" },
                { id: "file.new.fromExisting", label: "New from existing…", ariaLabel: "New from existing", iconId: "file" },
              ],
            },
            { id: "file.new.blankWorkbook", label: "Blank workbook", ariaLabel: "Blank workbook", iconId: "file" },
            {
              id: "file.new.templates",
              label: "Templates",
              ariaLabel: "Templates",
              iconId: "file",
              kind: "dropdown",
              menuItems: [
                { id: "file.new.templates.budget", label: "Budget", ariaLabel: "Budget template", iconId: "currency" },
                { id: "file.new.templates.invoice", label: "Invoice", ariaLabel: "Invoice template", iconId: "file" },
                { id: "file.new.templates.calendar", label: "Calendar", ariaLabel: "Calendar template", iconId: "calendar" },
                { id: "file.new.templates.more", label: "More…", ariaLabel: "More templates", iconId: "moreFormats" },
              ],
            },
          ],
        },
        {
          id: "file.info",
          label: "Info",
          buttons: [
            {
              id: "file.info.protectWorkbook",
              label: "Protect Workbook",
              ariaLabel: "Protect Workbook",
              iconId: "lock",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "file.info.protectWorkbook.encryptWithPassword", label: "Encrypt with Password…", ariaLabel: "Encrypt with Password", iconId: "lock" },
                { id: "file.info.protectWorkbook.protectCurrentSheet", label: "Protect Current Sheet…", ariaLabel: "Protect Current Sheet", iconId: "file" },
                { id: "file.info.protectWorkbook.protectWorkbookStructure", label: "Protect Workbook Structure…", ariaLabel: "Protect Workbook Structure", iconId: "settings" },
              ],
            },
            {
              id: "file.info.inspectWorkbook",
              label: "Inspect Workbook",
              ariaLabel: "Inspect Workbook",
              iconId: "search",
              kind: "dropdown",
              menuItems: [
                { id: "file.info.inspectWorkbook.documentInspector", label: "Document Inspector…", ariaLabel: "Document Inspector", iconId: "search" },
                { id: "file.info.inspectWorkbook.checkAccessibility", label: "Check Accessibility", ariaLabel: "Check Accessibility", iconId: "help" },
                { id: "file.info.inspectWorkbook.checkCompatibility", label: "Check Compatibility", ariaLabel: "Check Compatibility", iconId: "check" },
              ],
            },
            {
              id: "file.info.manageWorkbook",
              label: "Manage Workbook",
              ariaLabel: "Manage Workbook",
              iconId: "folderOpen",
              kind: "dropdown",
              menuItems: [
                 { id: "file.info.manageWorkbook.recoverUnsaved", label: "Recover Unsaved Workbooks…", ariaLabel: "Recover Unsaved Workbooks", iconId: "clock" },
                 { id: "file.info.manageWorkbook.versions", label: "Version History", ariaLabel: "Version History" },
                 { id: "file.info.manageWorkbook.branches", label: "Branches", ariaLabel: "Branches" },
                 { id: "file.info.manageWorkbook.properties", label: "Properties", ariaLabel: "Properties", iconId: "settings" },
               ],
             },
           ],
         },
        {
          id: "file.open",
          label: "Open",
          buttons: [
            { id: "file.open.open", label: "Open", ariaLabel: "Open", iconId: "folderOpen", size: "large" },
            {
              id: "file.open.recent",
              label: "Recent",
              ariaLabel: "Recent",
              iconId: "clock",
              kind: "dropdown",
              menuItems: [
                { id: "file.open.recent.book1", label: "Book1.xlsx", ariaLabel: "Open Book1", iconId: "file" },
                { id: "file.open.recent.budget", label: "Budget.xlsx", ariaLabel: "Open Budget", iconId: "file" },
                { id: "file.open.recent.forecast", label: "Forecast.xlsx", ariaLabel: "Open Forecast", iconId: "file" },
                { id: "file.open.recent.more", label: "More…", ariaLabel: "More recent files", iconId: "moreFormats" },
              ],
            },
            {
              id: "file.open.pinned",
              label: "Pinned",
              ariaLabel: "Pinned",
              iconId: "pin",
              kind: "dropdown",
              menuItems: [
                { id: "file.open.pinned.q4", label: "Q4 Forecast.xlsx", ariaLabel: "Open Q4 Forecast", iconId: "pin" },
                { id: "file.open.pinned.kpis", label: "KPIs.xlsx", ariaLabel: "Open KPIs", iconId: "pin" },
              ],
            },
          ],
        },
        {
          id: "file.save",
          label: "Save",
          buttons: [
            { id: "file.save.save", label: "Save", ariaLabel: "Save", iconId: "save", size: "large", testId: "ribbon-save" },
            {
              id: "file.save.saveAs",
              label: "Save As",
              ariaLabel: "Save As",
              iconId: "edit",
              kind: "dropdown",
              menuItems: [
                { id: "file.save.saveAs", label: "Save As…", ariaLabel: "Save As", iconId: "edit" },
                { id: "file.save.saveAs.copy", label: "Save a Copy…", ariaLabel: "Save a Copy", iconId: "file" },
                { id: "file.save.saveAs.download", label: "Download a Copy", ariaLabel: "Download a Copy", iconId: "arrowDown" },
              ],
            },
            { id: "file.save.autoSave", label: "AutoSave", ariaLabel: "AutoSave", iconId: "clock", kind: "toggle", defaultPressed: false },
          ],
        },
        {
          id: "file.export",
          label: "Export",
          buttons: [
            {
              id: "file.export.export",
              label: "Export",
              ariaLabel: "Export",
              iconId: "upload",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "file.export.export.pdf", label: "PDF", ariaLabel: "Export to PDF", iconId: "file" },
                { id: "file.export.export.csv", label: "CSV", ariaLabel: "Export to CSV", iconId: "file" },
                { id: "file.export.export.xlsx", label: "Excel Workbook", ariaLabel: "Export to Excel workbook", iconId: "chart" },
              ],
            },
            { id: "file.export.createPdf", label: "Create PDF/XPS", ariaLabel: "Create PDF or XPS", iconId: "file" },
            {
              id: "file.export.changeFileType",
              label: "Change File Type",
              ariaLabel: "Change File Type",
              iconId: "refresh",
              kind: "dropdown",
              menuItems: [
                { id: "file.export.changeFileType.xlsx", label: "Excel Workbook (*.xlsx)", ariaLabel: "Excel Workbook", iconId: "chart" },
                { id: "file.export.changeFileType.csv", label: "CSV (Comma delimited) (*.csv)", ariaLabel: "CSV", iconId: "file" },
                { id: "file.export.changeFileType.tsv", label: "TSV (Tab delimited) (*.tsv)", ariaLabel: "TSV", iconId: "return" },
                { id: "file.export.changeFileType.pdf", label: "PDF (*.pdf)", ariaLabel: "PDF", iconId: "file" },
              ],
            },
          ],
        },
        {
          id: "file.print",
          label: "Print",
          buttons: [
            { id: "file.print.print", label: "Print", ariaLabel: "Print", iconId: "print", size: "large", testId: "ribbon-print" },
            { id: "file.print.printPreview", label: "Print Preview", ariaLabel: "Print Preview", iconId: "eye" },
            {
              id: "file.print.pageSetup",
              label: "Page Setup",
              ariaLabel: "Page Setup",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "file.print.pageSetup", label: "Page Setup…", ariaLabel: "Page Setup", iconId: "settings" },
                { id: "file.print.pageSetup.printTitles", label: "Print Titles…", ariaLabel: "Print Titles", iconId: "tag" },
                { id: "file.print.pageSetup.margins", label: "Margins", ariaLabel: "Margins", iconId: "settings" },
              ],
            },
          ],
        },
        {
          id: "file.share",
          label: "Share",
          buttons: [
            { id: "file.share.share", label: "Share", ariaLabel: "Share", iconId: "link", size: "large" },
            {
              id: "file.share.email",
              label: "Email",
              ariaLabel: "Email",
              iconId: "mail",
              kind: "dropdown",
              menuItems: [
                { id: "file.share.email.attachment", label: "Send as Attachment", ariaLabel: "Send as Attachment", iconId: "clipboardPane" },
                { id: "file.share.email.link", label: "Send Link", ariaLabel: "Send Link", iconId: "link" },
              ],
            },
            { id: "file.share.presentOnline", label: "Present Online", ariaLabel: "Present Online", iconId: "globe" },
          ],
        },
        {
          id: "file.options",
          label: "Options",
          buttons: [
            { id: "file.options.options", label: "Options", ariaLabel: "Options", iconId: "settings", size: "large" },
            { id: "file.options.account", label: "Account", ariaLabel: "Account", iconId: "user" },
            { id: "file.options.close", label: "Close", ariaLabel: "Close", iconId: "close", testId: "ribbon-close" },
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
              iconId: "paste",
              kind: "dropdown",
              size: "large",
              testId: "ribbon-paste",
              menuItems: [
                { id: "home.clipboard.paste.default", label: "Paste", ariaLabel: "Paste", iconId: "paste" },
                { id: "home.clipboard.paste.values", label: "Paste Values", ariaLabel: "Paste Values", icon: "123" },
                { id: "home.clipboard.paste.formulas", label: "Paste Formulas", ariaLabel: "Paste Formulas", iconId: "function" },
                { id: "home.clipboard.paste.formats", label: "Paste Formatting", ariaLabel: "Paste Formatting", iconId: "palette" },
                { id: "home.clipboard.paste.transpose", label: "Transpose", ariaLabel: "Transpose", iconId: "refresh" },
              ],
            },
            {
              id: "home.clipboard.pasteSpecial",
              label: "Paste Special",
              ariaLabel: "Paste Special",
              iconId: "pasteSpecial",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.clipboard.pasteSpecial.dialog", label: "Paste Special…", ariaLabel: "Paste Special", iconId: "pasteSpecial" },
                { id: "home.clipboard.pasteSpecial.values", label: "Values", ariaLabel: "Paste Values", icon: "123" },
                { id: "home.clipboard.pasteSpecial.formulas", label: "Formulas", ariaLabel: "Paste Formulas", iconId: "function" },
                { id: "home.clipboard.pasteSpecial.formats", label: "Formats", ariaLabel: "Paste Formats", iconId: "palette" },
                { id: "home.clipboard.pasteSpecial.transpose", label: "Transpose", ariaLabel: "Transpose", iconId: "refresh" },
              ],
            },
            { id: "home.clipboard.cut", label: "Cut", ariaLabel: "Cut", iconId: "cut", size: "icon" },
            { id: "home.clipboard.copy", label: "Copy", ariaLabel: "Copy", iconId: "copy", size: "icon" },
            { id: "home.clipboard.formatPainter", label: "Format Painter", ariaLabel: "Format Painter", iconId: "formatPainter", size: "small" },
            {
              id: "home.clipboard.clipboardPane",
              label: "Clipboard",
              ariaLabel: "Open Clipboard",
              iconId: "clipboardPane",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.clipboard.clipboardPane.open", label: "Open Clipboard", ariaLabel: "Open Clipboard", iconId: "clipboardPane" },
                { id: "home.clipboard.clipboardPane.clearAll", label: "Clear All", ariaLabel: "Clear Clipboard", iconId: "clear" },
                { id: "home.clipboard.clipboardPane.options", label: "Options…", ariaLabel: "Clipboard Options", iconId: "settings" },
              ],
            },
          ],
        },
        {
          id: "home.debug.ai",
          label: "AI",
          buttons: [
            {
              id: "open-panel-ai-chat",
              label: "AI",
              ariaLabel: "Toggle AI panel",
              iconId: "sparkles",
              testId: "open-panel-ai-chat",
              size: "icon",
            },
            {
              id: "open-inline-ai-edit",
              label: "Inline Edit",
              ariaLabel: "Inline AI Edit",
              iconId: "sparkles",
              testId: "open-inline-ai-edit",
              size: "icon",
            },
            {
              id: "open-ai-panel",
              label: "AI (legacy)",
              ariaLabel: "Toggle AI panel",
              iconId: "sparkles",
              testId: "open-ai-panel",
              size: "icon",
            },
          ],
        },
        {
          id: "home.find",
          label: "Find",
          buttons: [
            {
              id: "home.editing.findSelect",
              label: "Find & Select",
              ariaLabel: "Find and Select",
              iconId: "find",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.findSelect.find", label: "Find", ariaLabel: "Find", iconId: "find", testId: "ribbon-find" },
                { id: "home.editing.findSelect.replace", label: "Replace", ariaLabel: "Replace", iconId: "replace", testId: "ribbon-replace" },
                { id: "home.editing.findSelect.goTo", label: "Go To", ariaLabel: "Go To", iconId: "goTo", testId: "ribbon-goto" },
              ],
            },
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
              iconId: "fontSize",
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
            { id: "home.font.increaseFont", label: "Grow Font", ariaLabel: "Increase Font Size", iconId: "increaseFont", size: "icon" },
            { id: "home.font.decreaseFont", label: "Shrink Font", ariaLabel: "Decrease Font Size", iconId: "decreaseFont", size: "icon" },
            { id: "home.font.bold", label: "Bold", ariaLabel: "Bold", iconId: "bold", kind: "toggle", size: "icon" },
            { id: "home.font.italic", label: "Italic", ariaLabel: "Italic", iconId: "italic", kind: "toggle", size: "icon" },
            { id: "home.font.underline", label: "Underline", ariaLabel: "Underline", iconId: "underline", kind: "toggle", size: "icon" },
            { id: "home.font.strikethrough", label: "Strike", ariaLabel: "Strikethrough", iconId: "strikethrough", kind: "toggle", size: "icon" },
            { id: "home.font.subscript", label: "Subscript", ariaLabel: "Subscript", iconId: "subscript", kind: "toggle", size: "icon" },
            { id: "home.font.superscript", label: "Superscript", ariaLabel: "Superscript", iconId: "superscript", kind: "toggle", size: "icon" },
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
                { id: "home.font.borders.thickBox", label: "Thick Box Border", ariaLabel: "Thick Box Border", iconId: "fillColor" },
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
              iconId: "fillColor",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                // Keep command ids aligned with desktop wiring (`apps/desktop/src/main.ts`)
                // so core formatting actions work out of the box.
                { id: "home.font.fillColor.none", label: "No Fill", ariaLabel: "No Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.lightGray", label: "Light Gray", ariaLabel: "Light Gray Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.yellow", label: "Yellow", ariaLabel: "Yellow Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.blue", label: "Blue", ariaLabel: "Blue Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.green", label: "Green", ariaLabel: "Green Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.red", label: "Red", ariaLabel: "Red Fill", iconId: "fillColor" },
                { id: "home.font.fillColor.moreColors", label: "More Colors…", ariaLabel: "More Fill Colors", iconId: "palette" },
              ],
            },
            {
              id: "home.font.fontColor",
              label: "Color",
              ariaLabel: "Font Color",
              iconId: "fontColor",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.font.fontColor.automatic", label: "Automatic", ariaLabel: "Automatic Font Color", icon: "A" },
                { id: "home.font.fontColor.black", label: "Black", ariaLabel: "Black Font Color", iconId: "fillColor" },
                { id: "home.font.fontColor.blue", label: "Blue", ariaLabel: "Blue Font Color", iconId: "fillColor" },
                { id: "home.font.fontColor.red", label: "Red", ariaLabel: "Red Font Color", iconId: "fillColor" },
                { id: "home.font.fontColor.green", label: "Green", ariaLabel: "Green Font Color", iconId: "fillColor" },
                { id: "home.font.fontColor.moreColors", label: "More Colors…", ariaLabel: "More Font Colors", iconId: "palette" },
              ],
            },
            {
              id: "home.font.clearFormatting",
              label: "Clear",
              ariaLabel: "Clear Formatting",
              iconId: "clearFormatting",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.font.clearFormatting.clearFormats", label: "Clear Formats", ariaLabel: "Clear Formats", iconId: "palette" },
                { id: "home.font.clearFormatting.clearContents", label: "Clear Contents", ariaLabel: "Clear Contents", icon: "⌫" },
                { id: "home.font.clearFormatting.clearAll", label: "Clear All", ariaLabel: "Clear All", iconId: "clear" },
              ],
            },
          ],
        },
        {
          id: "home.alignment",
          label: "Alignment",
          buttons: [
            { id: "home.alignment.topAlign", label: "Top", ariaLabel: "Top Align", iconId: "alignTop", size: "icon" },
            { id: "home.alignment.middleAlign", label: "Middle", ariaLabel: "Middle Align", iconId: "alignMiddle", size: "icon" },
            { id: "home.alignment.bottomAlign", label: "Bottom", ariaLabel: "Bottom Align", iconId: "alignBottom", size: "icon" },
            { id: "home.alignment.alignLeft", label: "Left", ariaLabel: "Align Left", iconId: "alignLeft", size: "icon" },
            { id: "home.alignment.center", label: "Center", ariaLabel: "Center", iconId: "alignCenter", size: "icon" },
            { id: "home.alignment.alignRight", label: "Right", ariaLabel: "Align Right", iconId: "alignRight", size: "icon" },
            {
              id: "home.alignment.orientation",
              label: "Orientation",
              ariaLabel: "Orientation",
              iconId: "orientation",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.alignment.orientation.angleCounterclockwise", label: "Angle Counterclockwise", ariaLabel: "Angle Counterclockwise", iconId: "orientation" },
                { id: "home.alignment.orientation.angleClockwise", label: "Angle Clockwise", ariaLabel: "Angle Clockwise", iconId: "orientation" },
                { id: "home.alignment.orientation.verticalText", label: "Vertical Text", ariaLabel: "Vertical Text", iconId: "arrowUpDown" },
                { id: "home.alignment.orientation.rotateUp", label: "Rotate Text Up", ariaLabel: "Rotate Text Up", iconId: "arrowUp" },
                { id: "home.alignment.orientation.rotateDown", label: "Rotate Text Down", ariaLabel: "Rotate Text Down", iconId: "arrowDown" },
                { id: "home.alignment.orientation.formatCellAlignment", label: "Format Cell Alignment…", ariaLabel: "Format Cell Alignment", iconId: "settings" },
              ],
            },
            { id: "home.alignment.wrapText", label: "Wrap Text", ariaLabel: "Wrap Text", iconId: "wrapText", kind: "toggle", size: "small" },
            {
              id: "home.alignment.mergeCenter",
              label: "Merge & Center",
              ariaLabel: "Merge and Center",
              iconId: "mergeCenter",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.alignment.mergeCenter.mergeCenter", label: "Merge & Center", ariaLabel: "Merge and Center", icon: "⊞" },
                { id: "home.alignment.mergeCenter.mergeAcross", label: "Merge Across", ariaLabel: "Merge Across", iconId: "arrowLeftRight" },
                { id: "home.alignment.mergeCenter.mergeCells", label: "Merge Cells", ariaLabel: "Merge Cells", icon: "▦" },
                { id: "home.alignment.mergeCenter.unmergeCells", label: "Unmerge Cells", ariaLabel: "Unmerge Cells", iconId: "close" },
              ],
            },
            { id: "home.alignment.increaseIndent", label: "Indent", ariaLabel: "Increase Indent", iconId: "increaseIndent", size: "icon" },
            { id: "home.alignment.decreaseIndent", label: "Outdent", ariaLabel: "Decrease Indent", iconId: "decreaseIndent", size: "icon" },
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
              iconId: "numberFormat",
              kind: "dropdown",
              size: "small",
              menuItems: [
                { id: "home.number.numberFormat.general", label: "General", ariaLabel: "General", icon: "123" },
                { id: "home.number.numberFormat.number", label: "Number", ariaLabel: "Number", icon: "0.00" },
                { id: "home.number.numberFormat.currency", label: "Currency", ariaLabel: "Currency", icon: "$" },
                { id: "home.number.numberFormat.accounting", label: "Accounting", ariaLabel: "Accounting", icon: "$" },
                { id: "home.number.numberFormat.shortDate", label: "Short Date", ariaLabel: "Short Date", iconId: "calendar" },
                { id: "home.number.numberFormat.longDate", label: "Long Date", ariaLabel: "Long Date", iconId: "calendar" },
                { id: "home.number.numberFormat.time", label: "Time", ariaLabel: "Time", iconId: "clock" },
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
              iconId: "currency",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.number.accounting.usd", label: "$ (Dollar)", ariaLabel: "Dollar", icon: "$" },
                { id: "home.number.accounting.eur", label: "€ (Euro)", ariaLabel: "Euro", icon: "€" },
                { id: "home.number.accounting.gbp", label: "£ (Pound)", ariaLabel: "Pound", icon: "£" },
                { id: "home.number.accounting.jpy", label: "¥ (Yen)", ariaLabel: "Yen", icon: "¥" },
              ],
            },
            { id: "home.number.percent", label: "Percent", ariaLabel: "Percent Style", iconId: "percent", size: "icon" },
            { id: "home.number.date", label: "Date", ariaLabel: "Date", iconId: "calendar", size: "icon" },
            { id: "home.number.comma", label: "Comma", ariaLabel: "Comma Style", iconId: "comma", size: "icon" },
            { id: "home.number.increaseDecimal", label: "Inc Decimal", ariaLabel: "Increase Decimal", iconId: "increaseDecimal", size: "icon" },
            { id: "home.number.decreaseDecimal", label: "Dec Decimal", ariaLabel: "Decrease Decimal", iconId: "decreaseDecimal", size: "icon" },
            {
              id: "home.number.moreFormats",
              label: "More",
              ariaLabel: "More Number Formats",
              iconId: "moreFormats",
              kind: "dropdown",
              size: "icon",
              menuItems: [
                { id: "home.number.moreFormats.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", iconId: "settings" },
                { id: "home.number.moreFormats.custom", label: "Custom…", ariaLabel: "Custom Number Format", iconId: "edit" },
              ],
            },
            { id: "home.number.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", iconId: "settings", size: "small", testId: "ribbon-format-cells" },
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
              iconId: "conditionalFormatting",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "home.styles.conditionalFormatting.highlightCellsRules", label: "Highlight Cells Rules", ariaLabel: "Highlight Cells Rules", iconId: "fillColor" },
                { id: "home.styles.conditionalFormatting.topBottomRules", label: "Top/Bottom Rules", ariaLabel: "Top or Bottom Rules", iconId: "arrowUp" },
                { id: "home.styles.conditionalFormatting.dataBars", label: "Data Bars", ariaLabel: "Data Bars", icon: "▮" },
                { id: "home.styles.conditionalFormatting.colorScales", label: "Color Scales", ariaLabel: "Color Scales", iconId: "palette" },
                { id: "home.styles.conditionalFormatting.iconSets", label: "Icon Sets", ariaLabel: "Icon Sets", iconId: "sparkles" },
                { id: "home.styles.conditionalFormatting.manageRules", label: "Manage Rules…", ariaLabel: "Manage Rules", iconId: "settings" },
                { id: "home.styles.conditionalFormatting.clearRules", label: "Clear Rules", ariaLabel: "Clear Rules", iconId: "close" },
              ],
            },
            {
              id: "home.styles.formatAsTable",
              label: "Format as Table",
              ariaLabel: "Format as Table",
              iconId: "formatAsTable",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "home.styles.formatAsTable.light", label: "Light", ariaLabel: "Light table styles", iconId: "file" },
                { id: "home.styles.formatAsTable.medium", label: "Medium", ariaLabel: "Medium table styles", iconId: "fillColor" },
                { id: "home.styles.formatAsTable.dark", label: "Dark", ariaLabel: "Dark table styles", iconId: "fillColor" },
                { id: "home.styles.formatAsTable.newStyle", label: "New Table Style…", ariaLabel: "New Table Style", iconId: "plus" },
              ],
            },
            {
              id: "home.styles.cellStyles",
              label: "Cell Styles",
              ariaLabel: "Cell Styles",
              iconId: "cellStyles",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "home.styles.cellStyles.goodBadNeutral", label: "Good, Bad, and Neutral", ariaLabel: "Good Bad and Neutral", iconId: "fillColor" },
                { id: "home.styles.cellStyles.dataModel", label: "Data Model", ariaLabel: "Data Model styles", iconId: "chart" },
                { id: "home.styles.cellStyles.titlesHeadings", label: "Titles and Headings", ariaLabel: "Titles and Headings", iconId: "tag" },
                { id: "home.styles.cellStyles.numberFormat", label: "Number Format", ariaLabel: "Number Format styles", icon: "123" },
                { id: "home.styles.cellStyles.newStyle", label: "New Cell Style…", ariaLabel: "New Cell Style", iconId: "plus" },
              ],
            },
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
              iconId: "insertCells",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.insert.insertCells", label: "Insert Cells…", ariaLabel: "Insert Cells", iconId: "insertCells" },
                { id: "home.cells.insert.insertSheetRows", label: "Insert Sheet Rows", ariaLabel: "Insert Sheet Rows", iconId: "insertRows" },
                { id: "home.cells.insert.insertSheetColumns", label: "Insert Sheet Columns", ariaLabel: "Insert Sheet Columns", iconId: "insertColumns" },
                { id: "home.cells.insert.insertSheet", label: "Insert Sheet", ariaLabel: "Insert Sheet", iconId: "insertSheet" },
              ],
            },
            {
              id: "home.cells.delete",
              label: "Delete",
              ariaLabel: "Delete Cells",
              iconId: "deleteCells",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.delete.deleteCells", label: "Delete Cells…", ariaLabel: "Delete Cells", iconId: "deleteCells" },
                { id: "home.cells.delete.deleteSheetRows", label: "Delete Sheet Rows", ariaLabel: "Delete Sheet Rows", iconId: "deleteCells" },
                { id: "home.cells.delete.deleteSheetColumns", label: "Delete Sheet Columns", ariaLabel: "Delete Sheet Columns", iconId: "deleteCells" },
                { id: "home.cells.delete.deleteSheet", label: "Delete Sheet", ariaLabel: "Delete Sheet", iconId: "deleteSheet" },
              ],
            },
            {
              id: "home.cells.format",
              label: "Format",
              ariaLabel: "Format Cells",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "home.cells.format.formatCells", label: "Format Cells…", ariaLabel: "Format Cells", iconId: "settings" },
                { id: "home.cells.format.rowHeight", label: "Row Height…", ariaLabel: "Row Height", iconId: "arrowUpDown" },
                { id: "home.cells.format.columnWidth", label: "Column Width…", ariaLabel: "Column Width", iconId: "arrowLeftRight" },
                { id: "home.cells.format.organizeSheets", label: "Organize Sheets", ariaLabel: "Organize Sheets", iconId: "folderOpen" },
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
              iconId: "autoSum",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.autoSum.sum", label: "Sum", ariaLabel: "Sum", iconId: "autoSum" },
                { id: "home.editing.autoSum.average", label: "Average", ariaLabel: "Average", iconId: "divide" },
                { id: "home.editing.autoSum.countNumbers", label: "Count Numbers", ariaLabel: "Count Numbers", iconId: "hash" },
                { id: "home.editing.autoSum.max", label: "Max", ariaLabel: "Max", iconId: "arrowUp" },
                { id: "home.editing.autoSum.min", label: "Min", ariaLabel: "Min", iconId: "arrowDown" },
                { id: "home.editing.autoSum.moreFunctions", label: "More Functions…", ariaLabel: "More Functions", iconId: "function" },
              ],
            },
            {
              id: "home.editing.fill",
              label: "Fill",
              ariaLabel: "Fill",
              iconId: "fillDown",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.fill.down", label: "Down", ariaLabel: "Fill Down", iconId: "arrowDown" },
                { id: "home.editing.fill.right", label: "Right", ariaLabel: "Fill Right", iconId: "arrowRight" },
                { id: "home.editing.fill.up", label: "Up", ariaLabel: "Fill Up", iconId: "arrowUp" },
                { id: "home.editing.fill.left", label: "Left", ariaLabel: "Fill Left", iconId: "arrowLeft" },
                { id: "home.editing.fill.series", label: "Series…", ariaLabel: "Series", iconId: "moreFormats" },
              ],
            },
            {
              id: "home.editing.clear",
              label: "Clear",
              ariaLabel: "Clear",
              iconId: "clear",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.clear.clearAll", label: "Clear All", ariaLabel: "Clear All", iconId: "clear" },
                { id: "home.editing.clear.clearFormats", label: "Clear Formats", ariaLabel: "Clear Formats", iconId: "palette" },
                { id: "home.editing.clear.clearContents", label: "Clear Contents", ariaLabel: "Clear Contents", iconId: "clear" },
                { id: "home.editing.clear.clearComments", label: "Clear Comments", ariaLabel: "Clear Comments", iconId: "comment" },
                { id: "home.editing.clear.clearHyperlinks", label: "Clear Hyperlinks", ariaLabel: "Clear Hyperlinks", iconId: "link" },
              ],
            },
            {
              id: "home.editing.sortFilter",
              label: "Sort & Filter",
              ariaLabel: "Sort and Filter",
              iconId: "sortFilter",
              kind: "dropdown",
              menuItems: [
                { id: "home.editing.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", iconId: "sort" },
                { id: "home.editing.sortFilter.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", iconId: "sort" },
                { id: "home.editing.sortFilter.customSort", label: "Custom Sort…", ariaLabel: "Custom Sort", iconId: "settings" },
                { id: "home.editing.sortFilter.filter", label: "Filter", ariaLabel: "Filter", iconId: "filter" },
                { id: "home.editing.sortFilter.clear", label: "Clear", ariaLabel: "Clear", iconId: "close" },
                { id: "home.editing.sortFilter.reapply", label: "Reapply", ariaLabel: "Reapply", iconId: "refresh" },
              ],
            },
          ],
        },
        {
          id: "home.debug.panels",
          label: "Panels",
          buttons: [
            {
              id: "open-panel-ai-audit",
              label: "Audit log (alt)",
              ariaLabel: "Toggle AI audit log panel",
              iconId: "file",
              testId: "open-panel-ai-audit",
              size: "icon",
            },
            {
              id: "open-ai-audit-panel",
              label: "Audit log (legacy)",
              ariaLabel: "Toggle AI audit log panel",
              iconId: "file",
              testId: "open-ai-audit-panel",
              size: "icon",
            },
            {
              id: "open-data-queries-panel",
              label: "Queries",
              ariaLabel: "Toggle Queries panel",
              iconId: "search",
              testId: "open-data-queries-panel",
              size: "icon",
            },
            {
              id: "open-macros-panel",
              label: "Macros",
              ariaLabel: "Toggle Macros panel",
              iconId: "file",
              testId: "open-macros-panel",
              size: "icon",
            },
            {
              id: "open-script-editor-panel",
              label: "Scripts",
              ariaLabel: "Toggle Script editor panel",
              iconId: "code",
              testId: "open-script-editor-panel",
              size: "icon",
            },
            {
              id: "open-python-panel",
              label: "Python",
              ariaLabel: "Toggle Python panel",
              iconId: "code",
              testId: "open-python-panel",
              size: "icon",
            },
            {
              id: "open-extensions-panel",
              label: "Extensions",
              ariaLabel: "Toggle Extensions panel",
              iconId: "puzzle",
              testId: "open-extensions-panel",
              size: "icon",
            },
            {
              id: "open-vba-migrate-panel",
              label: "Migrate Macros",
              ariaLabel: "Toggle VBA migrate panel",
              iconId: "settings",
              testId: "open-vba-migrate-panel",
              size: "icon",
            },
            {
              id: "open-comments-panel",
              label: "Comments",
              ariaLabel: "Toggle Comments panel",
              iconId: "comment",
              testId: "open-comments-panel",
              size: "icon",
            },
          ],
        },
        {
          id: "home.debug.auditing",
          label: "Auditing",
          buttons: [
            {
              id: "audit-precedents",
              label: "Trace precedents",
              ariaLabel: "Trace precedents",
              iconId: "arrowLeft",
              testId: "audit-precedents",
              size: "icon",
            },
            {
              id: "audit-dependents",
              label: "Trace dependents",
              ariaLabel: "Trace dependents",
              iconId: "arrowRight",
              testId: "audit-dependents",
              size: "icon",
            },
            {
              id: "audit-transitive",
              label: "Transitive",
              ariaLabel: "Toggle transitive auditing",
              iconId: "refresh",
              testId: "audit-transitive",
              size: "icon",
            },
          ],
        },
        {
          id: "home.debug.split",
          label: "Split view",
          buttons: [
            {
              id: "split-vertical",
              label: "Split vertical",
              ariaLabel: "Split vertically",
              iconId: "arrowLeftRight",
              testId: "split-vertical",
              size: "icon",
            },
            {
              id: "split-horizontal",
              label: "Split horizontal",
              ariaLabel: "Split horizontally",
              iconId: "arrowUpDown",
              testId: "split-horizontal",
              size: "icon",
            },
            { id: "split-none", label: "Unsplit", ariaLabel: "Remove split", iconId: "close", testId: "split-none", size: "icon" },
          ],
        },
        {
          id: "home.debug.freeze",
          label: "Freeze",
          buttons: [
            { id: "freeze-panes", label: "Freeze Panes", ariaLabel: "Freeze Panes", iconId: "lock", testId: "freeze-panes", size: "icon" },
            { id: "freeze-top-row", label: "Freeze Top Row", ariaLabel: "Freeze Top Row", iconId: "arrowUp", testId: "freeze-top-row", size: "icon" },
            { id: "freeze-first-column", label: "Freeze First Column", ariaLabel: "Freeze First Column", iconId: "arrowLeft", testId: "freeze-first-column", size: "icon" },
            { id: "unfreeze-panes", label: "Unfreeze Panes", ariaLabel: "Unfreeze Panes", iconId: "chart", testId: "unfreeze-panes", size: "icon" },
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
            {
              id: "insert.tables.pivotTable",
              label: "PivotTable",
              ariaLabel: "PivotTable",
              iconId: "chart",
              kind: "dropdown",
              size: "large",
              testId: "ribbon-insert-pivot-table",
              menuItems: [
                { id: "insert.tables.pivotTable", label: "PivotTable…", ariaLabel: "PivotTable", iconId: "chart" },
                { id: "insert.tables.pivotTable.fromTableRange", label: "From Table/Range…", ariaLabel: "PivotTable from Table or Range", icon: "▦" },
                { id: "insert.tables.pivotTable.fromExternal", label: "From External Data…", ariaLabel: "PivotTable from External Data", iconId: "globe" },
                { id: "insert.tables.pivotTable.fromDataModel", label: "From Data Model…", ariaLabel: "PivotTable from Data Model", iconId: "puzzle" },
              ],
            },
            {
              id: "insert.tables.recommendedPivotTables",
              label: "Recommended PivotTables",
              ariaLabel: "Recommended PivotTables",
              iconId: "sparkles",
              kind: "dropdown",
              menuItems: [
                {
                  id: "insert.tables.recommendedPivotTables",
                  label: "Recommended PivotTables…",
                  ariaLabel: "Recommended PivotTables",
                  iconId: "sparkles",
                },
              ],
            },
            { id: "insert.tables.table", label: "Table", ariaLabel: "Table", icon: "▦", size: "large" },
          ],
        },
        {
          id: "insert.pivotcharts",
          label: "PivotCharts",
          buttons: [
            { id: "insert.pivotcharts.pivotChart", label: "PivotChart", ariaLabel: "PivotChart", iconId: "chart", kind: "dropdown", size: "large" },
            { id: "insert.pivotcharts.recommendedPivotCharts", label: "Recommended PivotCharts", ariaLabel: "Recommended PivotCharts", iconId: "sparkles", kind: "dropdown" },
          ],
        },
        {
          id: "insert.illustrations",
          label: "Illustrations",
          buttons: [
            {
              id: "insert.illustrations.pictures",
              label: "Pictures",
              ariaLabel: "Pictures",
              iconId: "image",
              kind: "dropdown",
              menuItems: [
                { id: "insert.illustrations.pictures.thisDevice", label: "This Device…", ariaLabel: "Pictures from this device", iconId: "image" },
                { id: "insert.illustrations.pictures.stockImages", label: "Stock Images…", ariaLabel: "Stock Images", iconId: "sparkles" },
                { id: "insert.illustrations.pictures.onlinePictures", label: "Online Pictures…", ariaLabel: "Online Pictures", iconId: "globe" },
              ],
            },
            { id: "insert.illustrations.onlinePictures", label: "Online Pictures", ariaLabel: "Online Pictures", iconId: "globe", kind: "dropdown" },
            {
              id: "insert.illustrations.shapes",
              label: "Shapes",
              ariaLabel: "Shapes",
              iconId: "fillColor",
              kind: "dropdown",
              menuItems: [
                { id: "insert.illustrations.shapes.lines", label: "Lines", ariaLabel: "Lines", icon: "╱" },
                { id: "insert.illustrations.shapes.rectangles", label: "Rectangles", ariaLabel: "Rectangles", icon: "▭" },
                { id: "insert.illustrations.shapes.basicShapes", label: "Basic Shapes", ariaLabel: "Basic Shapes", iconId: "fillColor" },
                { id: "insert.illustrations.shapes.arrows", label: "Block Arrows", ariaLabel: "Block Arrows", iconId: "arrowRight" },
                { id: "insert.illustrations.shapes.flowchart", label: "Flowchart", ariaLabel: "Flowchart", iconId: "puzzle" },
                { id: "insert.illustrations.shapes.callouts", label: "Callouts", ariaLabel: "Callouts", iconId: "comment" },
              ],
            },
            { id: "insert.illustrations.icons", label: "Icons", ariaLabel: "Icons", iconId: "sparkles", kind: "dropdown" },
            { id: "insert.illustrations.smartArt", label: "SmartArt", ariaLabel: "SmartArt", iconId: "puzzle", kind: "dropdown" },
            { id: "insert.illustrations.screenshot", label: "Screenshot", ariaLabel: "Screenshot", iconId: "image", kind: "dropdown" },
          ],
        },
        {
          id: "insert.addins",
          label: "Add-ins",
          buttons: [
            { id: "insert.addins.getAddins", label: "Get Add-ins", ariaLabel: "Get Add-ins", iconId: "plus", kind: "dropdown" },
            { id: "insert.addins.myAddins", label: "My Add-ins", ariaLabel: "My Add-ins", iconId: "puzzle", kind: "dropdown" },
          ],
        },
        {
          id: "insert.charts",
          label: "Charts",
          buttons: [
            {
              id: "insert.charts.recommendedCharts",
              label: "Recommended Charts",
              ariaLabel: "Recommended Charts",
              iconId: "sparkles",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "insert.charts.recommendedCharts", label: "Recommended Charts…", ariaLabel: "Recommended Charts", iconId: "sparkles" },
                { id: "insert.charts.recommendedCharts.column", label: "Column", ariaLabel: "Recommended Column Charts", icon: "▮▮" },
                { id: "insert.charts.recommendedCharts.line", label: "Line", ariaLabel: "Recommended Line Charts", iconId: "chart" },
                { id: "insert.charts.recommendedCharts.pie", label: "Pie", ariaLabel: "Recommended Pie Charts", icon: "◔" },
                { id: "insert.charts.recommendedCharts.more", label: "More…", ariaLabel: "More chart recommendations", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.column",
              label: "Column",
              ariaLabel: "Insert Column or Bar Chart",
              icon: "▮▮",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.column.clusteredColumn", label: "Clustered Column", ariaLabel: "Clustered Column", icon: "▮▮" },
                { id: "insert.charts.column.stackedColumn", label: "Stacked Column", ariaLabel: "Stacked Column", icon: "▮▯" },
                { id: "insert.charts.column.stackedColumn100", label: "100% Stacked Column", ariaLabel: "100% Stacked Column", icon: "▮▮" },
                { id: "insert.charts.column.more", label: "More Column Charts…", ariaLabel: "More Column Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.line",
              label: "Line",
              ariaLabel: "Insert Line or Area Chart",
              iconId: "chart",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.line.line", label: "Line", ariaLabel: "Line", iconId: "chart" },
                { id: "insert.charts.line.lineWithMarkers", label: "Line with Markers", ariaLabel: "Line with Markers", icon: "•" },
                { id: "insert.charts.line.stackedArea", label: "Stacked Area", ariaLabel: "Stacked Area", iconId: "chart" },
                { id: "insert.charts.line.more", label: "More Line Charts…", ariaLabel: "More Line Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.pie",
              label: "Pie",
              ariaLabel: "Insert Pie or Doughnut Chart",
              icon: "◔",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.pie.pie", label: "Pie", ariaLabel: "Pie", icon: "◔" },
                { id: "insert.charts.pie.doughnut", label: "Doughnut", ariaLabel: "Doughnut", icon: "◑" },
                { id: "insert.charts.pie.more", label: "More Pie Charts…", ariaLabel: "More Pie Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.bar",
              label: "Bar",
              ariaLabel: "Insert Bar Chart",
              icon: "▭",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.bar.clusteredBar", label: "Clustered Bar", ariaLabel: "Clustered Bar", icon: "▭" },
                { id: "insert.charts.bar.stackedBar", label: "Stacked Bar", ariaLabel: "Stacked Bar", icon: "▭" },
                { id: "insert.charts.bar.more", label: "More Bar Charts…", ariaLabel: "More Bar Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.area",
              label: "Area",
              ariaLabel: "Insert Area Chart",
              iconId: "chart",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.area.area", label: "Area", ariaLabel: "Area", iconId: "chart" },
                { id: "insert.charts.area.stackedArea", label: "Stacked Area", ariaLabel: "Stacked Area", iconId: "chart" },
                { id: "insert.charts.area.more", label: "More Area Charts…", ariaLabel: "More Area Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.scatter",
              label: "Scatter",
              ariaLabel: "Insert Scatter (X, Y) Chart",
              iconId: "chart",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.scatter.scatter", label: "Scatter", ariaLabel: "Scatter", iconId: "chart" },
                { id: "insert.charts.scatter.smoothLines", label: "Scatter with Smooth Lines", ariaLabel: "Scatter with Smooth Lines", iconId: "chart" },
                { id: "insert.charts.scatter.more", label: "More Scatter Charts…", ariaLabel: "More Scatter Charts", iconId: "moreFormats" },
              ],
            },
            {
              id: "insert.charts.map",
              label: "Map",
              ariaLabel: "Insert Map Chart",
              iconId: "globe",
              kind: "dropdown",
              menuItems: [
                { id: "insert.charts.map.filledMap", label: "Filled Map", ariaLabel: "Filled Map", iconId: "globe" },
                { id: "insert.charts.map.more", label: "More Map Charts…", ariaLabel: "More Map Charts", iconId: "moreFormats" },
              ],
            },
            { id: "insert.charts.histogram", label: "Histogram", ariaLabel: "Insert Statistic Chart (Histogram, Pareto)", icon: "▁▃▆", kind: "dropdown" },
            { id: "insert.charts.waterfall", label: "Waterfall", ariaLabel: "Insert Waterfall Chart", iconId: "chart", kind: "dropdown" },
            { id: "insert.charts.treemap", label: "Treemap", ariaLabel: "Insert Hierarchy Chart (Treemap)", iconId: "puzzle", kind: "dropdown" },
            { id: "insert.charts.sunburst", label: "Sunburst", ariaLabel: "Insert Hierarchy Chart (Sunburst)", iconId: "chart", kind: "dropdown" },
            { id: "insert.charts.funnel", label: "Funnel", ariaLabel: "Insert Funnel Chart", iconId: "chart", kind: "dropdown" },
            { id: "insert.charts.boxWhisker", label: "Box & Whisker", ariaLabel: "Insert Box and Whisker Chart", icon: "▣", kind: "dropdown" },
            { id: "insert.charts.radar", label: "Radar", ariaLabel: "Insert Radar Chart", iconId: "globe", kind: "dropdown" },
            { id: "insert.charts.surface", label: "Surface", ariaLabel: "Insert Surface Chart", iconId: "chart", kind: "dropdown" },
            { id: "insert.charts.stock", label: "Stock", ariaLabel: "Insert Stock Chart", iconId: "chart", kind: "dropdown" },
            { id: "insert.charts.combo", label: "Combo", ariaLabel: "Insert Combo Chart", iconId: "shuffle", kind: "dropdown" },
            { id: "insert.charts.pivotChart", label: "PivotChart", ariaLabel: "PivotChart", iconId: "chart", kind: "dropdown" },
          ],
        },
        {
          id: "insert.tours",
          label: "Tours",
          buttons: [
            { id: "insert.tours.3dMap", label: "3D Map", ariaLabel: "3D Map", iconId: "globe", kind: "dropdown", size: "large" },
            { id: "insert.tours.launchTour", label: "Launch Tour", ariaLabel: "Launch Tour", iconId: "play", kind: "dropdown" },
          ],
        },
        {
          id: "insert.sparklines",
          label: "Sparklines",
          buttons: [
            {
              id: "insert.sparklines.line",
              label: "Line",
              ariaLabel: "Insert Line Sparkline",
              icon: "╱",
              kind: "dropdown",
              menuItems: [{ id: "insert.sparklines.line", label: "Line Sparkline…", ariaLabel: "Line Sparkline", icon: "╱" }],
            },
            {
              id: "insert.sparklines.column",
              label: "Column",
              ariaLabel: "Insert Column Sparkline",
              icon: "▮",
              kind: "dropdown",
              menuItems: [{ id: "insert.sparklines.column", label: "Column Sparkline…", ariaLabel: "Column Sparkline", icon: "▮" }],
            },
            {
              id: "insert.sparklines.winLoss",
              label: "Win/Loss",
              ariaLabel: "Insert Win/Loss Sparkline",
              icon: "±",
              kind: "dropdown",
              menuItems: [{ id: "insert.sparklines.winLoss", label: "Win/Loss Sparkline…", ariaLabel: "Win/Loss Sparkline", icon: "±" }],
            },
          ],
        },
        {
          id: "insert.filters",
          label: "Filters",
          buttons: [
            {
              id: "insert.filters.slicer",
              label: "Slicer",
              ariaLabel: "Insert Slicer",
              iconId: "cut",
              kind: "dropdown",
              menuItems: [
                { id: "insert.filters.slicer", label: "Slicer…", ariaLabel: "Insert Slicer", iconId: "cut" },
                { id: "insert.filters.slicer.reportConnections", label: "Report Connections…", ariaLabel: "Report Connections", iconId: "link" },
              ],
            },
            {
              id: "insert.filters.timeline",
              label: "Timeline",
              ariaLabel: "Insert Timeline",
              iconId: "clock",
              kind: "dropdown",
              menuItems: [{ id: "insert.filters.timeline", label: "Timeline…", ariaLabel: "Insert Timeline", iconId: "clock" }],
            },
          ],
        },
        {
          id: "insert.links",
          label: "Links",
          buttons: [{ id: "insert.links.link", label: "Link", ariaLabel: "Insert Link", iconId: "link", kind: "dropdown", size: "large" }],
        },
        {
          id: "insert.comments",
          label: "Comments",
          buttons: [
            { id: "insert.comments.comment", label: "Comment", ariaLabel: "Insert Comment", iconId: "comment", kind: "dropdown", size: "large" },
            { id: "insert.comments.note", label: "Note", ariaLabel: "Insert Note", iconId: "file", kind: "dropdown" },
          ],
        },
        {
          id: "insert.text",
          label: "Text",
          buttons: [
            { id: "insert.text.textBox", label: "Text Box", ariaLabel: "Insert Text Box", iconId: "edit", kind: "dropdown" },
            { id: "insert.text.headerFooter", label: "Header & Footer", ariaLabel: "Header and Footer", iconId: "file", kind: "dropdown" },
            { id: "insert.text.wordArt", label: "WordArt", ariaLabel: "WordArt", icon: "𝒜", kind: "dropdown" },
            { id: "insert.text.signatureLine", label: "Signature Line", ariaLabel: "Signature Line", iconId: "edit", kind: "dropdown" },
            { id: "insert.text.object", label: "Object", ariaLabel: "Object", iconId: "file", kind: "dropdown" },
          ],
        },
        {
          id: "insert.equations",
          label: "Equations",
          buttons: [
            { id: "insert.equations.equation", label: "Equation", ariaLabel: "Insert Equation", icon: "∑", kind: "dropdown", size: "large" },
            { id: "insert.equations.inkEquation", label: "Ink Equation", ariaLabel: "Ink Equation", iconId: "edit", kind: "dropdown" },
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
            {
              id: "pageLayout.themes.themes",
              label: "Themes",
              ariaLabel: "Themes",
              iconId: "settings",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "pageLayout.themes.themes.office", label: "Office", ariaLabel: "Office Theme", iconId: "settings" },
                { id: "pageLayout.themes.themes.integral", label: "Integral", ariaLabel: "Integral Theme", iconId: "settings" },
                { id: "pageLayout.themes.themes.facet", label: "Facet", ariaLabel: "Facet Theme", iconId: "settings" },
                { id: "pageLayout.themes.themes.customize", label: "Customize…", ariaLabel: "Customize Theme", iconId: "settings" },
              ],
            },
            {
              id: "pageLayout.themes.colors",
              label: "Colors",
              ariaLabel: "Colors",
              iconId: "palette",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.themes.colors.office", label: "Office", ariaLabel: "Office Colors", iconId: "palette" },
                { id: "pageLayout.themes.colors.colorful", label: "Colorful", ariaLabel: "Colorful Colors", iconId: "palette" },
                { id: "pageLayout.themes.colors.customize", label: "Customize Colors…", ariaLabel: "Customize Colors", iconId: "settings" },
              ],
            },
            {
              id: "pageLayout.themes.fonts",
              label: "Fonts",
              ariaLabel: "Fonts",
              iconId: "numberFormat",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.themes.fonts.office", label: "Office", ariaLabel: "Office Fonts", iconId: "numberFormat" },
                { id: "pageLayout.themes.fonts.aptos", label: "Aptos", ariaLabel: "Aptos Fonts", icon: "A" },
                { id: "pageLayout.themes.fonts.customize", label: "Customize Fonts…", ariaLabel: "Customize Fonts", iconId: "settings" },
              ],
            },
            {
              id: "pageLayout.themes.effects",
              label: "Effects",
              ariaLabel: "Effects",
              iconId: "sparkles",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.themes.effects.subtle", label: "Subtle", ariaLabel: "Subtle Effects", iconId: "sparkles" },
                { id: "pageLayout.themes.effects.moderate", label: "Moderate", ariaLabel: "Moderate Effects", iconId: "sparkles" },
                { id: "pageLayout.themes.effects.intense", label: "Intense", ariaLabel: "Intense Effects", iconId: "sparkles" },
              ],
            },
          ],
        },
        {
          id: "pageLayout.pageSetup",
          label: "Page Setup",
          buttons: [
            { id: "pageLayout.pageSetup.pageSetupDialog", label: "Page Setup…", ariaLabel: "Page Setup", iconId: "settings", size: "large", testId: "ribbon-page-setup" },
            {
              id: "pageLayout.pageSetup.margins",
              label: "Margins",
              ariaLabel: "Margins",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.margins.normal", label: "Normal", ariaLabel: "Normal Margins", iconId: "settings" },
                { id: "pageLayout.pageSetup.margins.wide", label: "Wide", ariaLabel: "Wide Margins", iconId: "arrowLeftRight" },
                { id: "pageLayout.pageSetup.margins.narrow", label: "Narrow", ariaLabel: "Narrow Margins", iconId: "arrowUpDown" },
                { id: "pageLayout.pageSetup.margins.custom", label: "Custom Margins…", ariaLabel: "Custom Margins", iconId: "settings" },
              ],
            },
            {
              id: "pageLayout.pageSetup.orientation",
              label: "Orientation",
              ariaLabel: "Orientation",
              iconId: "arrowLeftRight",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.orientation.portrait", label: "Portrait", ariaLabel: "Portrait", iconId: "file" },
                { id: "pageLayout.pageSetup.orientation.landscape", label: "Landscape", ariaLabel: "Landscape", icon: "▭" },
              ],
            },
            {
              id: "pageLayout.pageSetup.size",
              label: "Size",
              ariaLabel: "Size",
              iconId: "file",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.size.letter", label: "Letter", ariaLabel: "Letter", iconId: "file" },
                { id: "pageLayout.pageSetup.size.a4", label: "A4", ariaLabel: "A4", iconId: "file" },
                { id: "pageLayout.pageSetup.size.more", label: "More Paper Sizes…", ariaLabel: "More Paper Sizes", iconId: "moreFormats" },
              ],
            },
            {
              id: "pageLayout.pageSetup.printArea",
              label: "Print Area",
              ariaLabel: "Print Area",
              iconId: "print",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.printArea.set", label: "Set Print Area", ariaLabel: "Set Print Area", iconId: "print" },
                { id: "pageLayout.pageSetup.printArea.clear", label: "Clear Print Area", ariaLabel: "Clear Print Area", iconId: "close" },
                { id: "pageLayout.pageSetup.printArea.addTo", label: "Add to Print Area", ariaLabel: "Add to Print Area", iconId: "plus" },
              ],
            },
            {
              id: "pageLayout.pageSetup.breaks",
              label: "Breaks",
              ariaLabel: "Breaks",
              iconId: "pageBreak",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.breaks.insertPageBreak", label: "Insert Page Break", ariaLabel: "Insert Page Break", iconId: "pageBreak" },
                { id: "pageLayout.pageSetup.breaks.removePageBreak", label: "Remove Page Break", ariaLabel: "Remove Page Break", iconId: "close" },
                { id: "pageLayout.pageSetup.breaks.resetAll", label: "Reset All Page Breaks", ariaLabel: "Reset All Page Breaks", iconId: "undo" },
              ],
            },
            {
              id: "pageLayout.pageSetup.background",
              label: "Background",
              ariaLabel: "Background",
              iconId: "image",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.background.background", label: "Background…", ariaLabel: "Background", iconId: "image" },
                { id: "pageLayout.pageSetup.background.delete", label: "Delete Background", ariaLabel: "Delete Background", iconId: "trash" },
              ],
            },
            {
              id: "pageLayout.pageSetup.printTitles",
              label: "Print Titles",
              ariaLabel: "Print Titles",
              iconId: "tag",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.pageSetup.printTitles.printTitles", label: "Print Titles…", ariaLabel: "Print Titles", iconId: "tag" },
              ],
            },
          ],
        },
        {
          id: "pageLayout.printArea",
          label: "Print Area",
          buttons: [
            { id: "pageLayout.printArea.setPrintArea", label: "Set Print Area", ariaLabel: "Set Print Area", iconId: "print", testId: "ribbon-set-print-area" },
            { id: "pageLayout.printArea.clearPrintArea", label: "Clear Print Area", ariaLabel: "Clear Print Area", iconId: "close", testId: "ribbon-clear-print-area" },
          ],
        },
        {
          id: "pageLayout.export",
          label: "Export",
          buttons: [
            { id: "pageLayout.export.exportPdf", label: "Export to PDF", ariaLabel: "Export to PDF", iconId: "file", testId: "ribbon-export-pdf" },
          ],
        },
        {
          id: "pageLayout.scaleToFit",
          label: "Scale to Fit",
          buttons: [
            {
              id: "pageLayout.scaleToFit.width",
              label: "Width",
              ariaLabel: "Width",
              iconId: "arrowLeftRight",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.scaleToFit.width.automatic", label: "Automatic", ariaLabel: "Automatic width", icon: "∞" },
                { id: "pageLayout.scaleToFit.width.1page", label: "1 page", ariaLabel: "Fit to 1 page wide", icon: "1" },
                { id: "pageLayout.scaleToFit.width.2pages", label: "2 pages", ariaLabel: "Fit to 2 pages wide", icon: "2" },
              ],
            },
            {
              id: "pageLayout.scaleToFit.height",
              label: "Height",
              ariaLabel: "Height",
              iconId: "arrowUpDown",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.scaleToFit.height.automatic", label: "Automatic", ariaLabel: "Automatic height", icon: "∞" },
                { id: "pageLayout.scaleToFit.height.1page", label: "1 page", ariaLabel: "Fit to 1 page tall", icon: "1" },
                { id: "pageLayout.scaleToFit.height.2pages", label: "2 pages", ariaLabel: "Fit to 2 pages tall", icon: "2" },
              ],
            },
            {
              id: "pageLayout.scaleToFit.scale",
              label: "Scale",
              ariaLabel: "Scale",
              iconId: "search",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.scaleToFit.scale.100", label: "100%", ariaLabel: "Scale 100%", icon: "100%" },
                { id: "pageLayout.scaleToFit.scale.90", label: "90%", ariaLabel: "Scale 90%", icon: "90%" },
                { id: "pageLayout.scaleToFit.scale.80", label: "80%", ariaLabel: "Scale 80%", icon: "80%" },
                { id: "pageLayout.scaleToFit.scale.70", label: "70%", ariaLabel: "Scale 70%", icon: "70%" },
                { id: "pageLayout.scaleToFit.scale.more", label: "Custom…", ariaLabel: "Custom scale", iconId: "settings" },
              ],
            },
          ],
        },
        {
          id: "pageLayout.sheetOptions",
          label: "Sheet Options",
          buttons: [
            { id: "pageLayout.sheetOptions.gridlinesView", label: "Gridlines View", ariaLabel: "View Gridlines", icon: "▦", kind: "toggle", size: "small", defaultPressed: true },
            { id: "pageLayout.sheetOptions.gridlinesPrint", label: "Gridlines Print", ariaLabel: "Print Gridlines", iconId: "print", kind: "toggle", size: "small", defaultPressed: false },
            { id: "pageLayout.sheetOptions.headingsView", label: "Headings View", ariaLabel: "View Headings", icon: "A1", kind: "toggle", size: "small", defaultPressed: true },
            { id: "pageLayout.sheetOptions.headingsPrint", label: "Headings Print", ariaLabel: "Print Headings", iconId: "print", kind: "toggle", size: "small", defaultPressed: false },
          ],
        },
        {
          id: "pageLayout.arrange",
          label: "Arrange",
          buttons: [
            { id: "pageLayout.arrange.bringForward", label: "Bring Forward", ariaLabel: "Bring Forward", iconId: "arrowUp", kind: "dropdown" },
            { id: "pageLayout.arrange.sendBackward", label: "Send Backward", ariaLabel: "Send Backward", iconId: "arrowDown", kind: "dropdown" },
            { id: "pageLayout.arrange.selectionPane", label: "Selection Pane", ariaLabel: "Selection Pane", iconId: "paste" },
            {
              id: "pageLayout.arrange.align",
              label: "Align",
              ariaLabel: "Align",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.arrange.align.alignLeft", label: "Align Left", ariaLabel: "Align Left", iconId: "arrowLeft" },
                { id: "pageLayout.arrange.align.alignCenter", label: "Align Center", ariaLabel: "Align Center", iconId: "arrowLeftRight" },
                { id: "pageLayout.arrange.align.alignRight", label: "Align Right", ariaLabel: "Align Right", iconId: "arrowRight" },
                { id: "pageLayout.arrange.align.alignTop", label: "Align Top", ariaLabel: "Align Top", iconId: "arrowUp" },
                { id: "pageLayout.arrange.align.alignMiddle", label: "Align Middle", ariaLabel: "Align Middle", iconId: "arrowUpDown" },
                { id: "pageLayout.arrange.align.alignBottom", label: "Align Bottom", ariaLabel: "Align Bottom", iconId: "arrowDown" },
              ],
            },
            {
              id: "pageLayout.arrange.group",
              label: "Group",
              ariaLabel: "Group",
              iconId: "link",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.arrange.group.group", label: "Group", ariaLabel: "Group", iconId: "link" },
                { id: "pageLayout.arrange.group.ungroup", label: "Ungroup", ariaLabel: "Ungroup", iconId: "close" },
                { id: "pageLayout.arrange.group.regroup", label: "Regroup", ariaLabel: "Regroup", iconId: "redo" },
              ],
            },
            {
              id: "pageLayout.arrange.rotate",
              label: "Rotate",
              ariaLabel: "Rotate",
              iconId: "redo",
              kind: "dropdown",
              menuItems: [
                { id: "pageLayout.arrange.rotate.rotateRight90", label: "Rotate Right 90°", ariaLabel: "Rotate Right 90 degrees", iconId: "redo" },
                { id: "pageLayout.arrange.rotate.rotateLeft90", label: "Rotate Left 90°", ariaLabel: "Rotate Left 90 degrees", iconId: "undo" },
                { id: "pageLayout.arrange.rotate.flipVertical", label: "Flip Vertical", ariaLabel: "Flip Vertical", iconId: "arrowUpDown" },
                { id: "pageLayout.arrange.rotate.flipHorizontal", label: "Flip Horizontal", ariaLabel: "Flip Horizontal", iconId: "arrowLeftRight" },
              ],
            },
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
             { id: "formulas.functionLibrary.insertFunction", label: "Insert Function", ariaLabel: "Insert Function", iconId: "function", kind: "dropdown", size: "large" },
             {
               id: "formulas.functionLibrary.autoSum",
               label: "AutoSum",
               ariaLabel: "AutoSum",
               iconId: "autoSum",
               kind: "dropdown",
               menuItems: [
                 { id: "formulas.functionLibrary.autoSum.sum", label: "Sum", ariaLabel: "Sum", iconId: "autoSum" },
                 { id: "formulas.functionLibrary.autoSum.average", label: "Average", ariaLabel: "Average", iconId: "divide" },
                 { id: "formulas.functionLibrary.autoSum.countNumbers", label: "Count Numbers", ariaLabel: "Count Numbers", iconId: "hash" },
                 { id: "formulas.functionLibrary.autoSum.max", label: "Max", ariaLabel: "Max", iconId: "arrowUp" },
                 { id: "formulas.functionLibrary.autoSum.min", label: "Min", ariaLabel: "Min", iconId: "arrowDown" },
                 { id: "formulas.functionLibrary.autoSum.moreFunctions", label: "More Functions…", ariaLabel: "More Functions", iconId: "function" },
               ],
             },
            { id: "formulas.functionLibrary.recentlyUsed", label: "Recently Used", ariaLabel: "Recently Used", iconId: "clock", kind: "dropdown" },
            { id: "formulas.functionLibrary.financial", label: "Financial", ariaLabel: "Financial", icon: "$", kind: "dropdown" },
            { id: "formulas.functionLibrary.logical", label: "Logical", ariaLabel: "Logical", icon: "∧", kind: "dropdown" },
            { id: "formulas.functionLibrary.text", label: "Text", ariaLabel: "Text", icon: "Aa", kind: "dropdown" },
            { id: "formulas.functionLibrary.dateTime", label: "Date & Time", ariaLabel: "Date and Time", iconId: "calendar", kind: "dropdown" },
            { id: "formulas.functionLibrary.lookupReference", label: "Lookup & Reference", ariaLabel: "Lookup and Reference", iconId: "search", kind: "dropdown" },
            { id: "formulas.functionLibrary.mathTrig", label: "Math & Trig", ariaLabel: "Math and Trig", icon: "π", kind: "dropdown" },
            { id: "formulas.functionLibrary.moreFunctions", label: "More Functions", ariaLabel: "More Functions", iconId: "plus", kind: "dropdown" },
          ],
        },
        {
          id: "formulas.definedNames",
          label: "Defined Names",
          buttons: [
            { id: "formulas.definedNames.nameManager", label: "Name Manager", ariaLabel: "Name Manager", iconId: "tag", kind: "dropdown", size: "large" },
            { id: "formulas.definedNames.defineName", label: "Define Name", ariaLabel: "Define Name", iconId: "plus", kind: "dropdown" },
            { id: "formulas.definedNames.useInFormula", label: "Use in Formula", ariaLabel: "Use in Formula", iconId: "function", kind: "dropdown" },
            { id: "formulas.definedNames.createFromSelection", label: "Create from Selection", ariaLabel: "Create from Selection", icon: "▦", kind: "dropdown" },
          ],
        },
        {
          id: "formulas.formulaAuditing",
          label: "Formula Auditing",
          buttons: [
            { id: "formulas.formulaAuditing.tracePrecedents", label: "Trace Precedents", ariaLabel: "Trace Precedents", iconId: "arrowLeft", size: "small" },
            { id: "formulas.formulaAuditing.traceDependents", label: "Trace Dependents", ariaLabel: "Trace Dependents", iconId: "arrowRight", size: "small" },
            { id: "formulas.formulaAuditing.removeArrows", label: "Remove Arrows", ariaLabel: "Remove Arrows", iconId: "close", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.showFormulas", label: "Show Formulas", ariaLabel: "Show Formulas", iconId: "function", kind: "toggle", size: "small" },
            { id: "formulas.formulaAuditing.errorChecking", label: "Error Checking", ariaLabel: "Error Checking", iconId: "warning", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.evaluateFormula", label: "Evaluate Formula", ariaLabel: "Evaluate Formula", iconId: "autoSum", kind: "dropdown", size: "small" },
            { id: "formulas.formulaAuditing.watchWindow", label: "Watch Window", ariaLabel: "Watch Window", iconId: "eye", kind: "dropdown", size: "small" },
          ],
        },
        {
          id: "formulas.calculation",
          label: "Calculation",
          buttons: [
            { id: "formulas.calculation.calculationOptions", label: "Calculation Options", ariaLabel: "Calculation Options", iconId: "settings", kind: "dropdown", size: "large" },
            { id: "formulas.calculation.calculateNow", label: "Calculate Now", ariaLabel: "Calculate Now", iconId: "refresh", size: "small" },
            { id: "formulas.calculation.calculateSheet", label: "Calculate Sheet", ariaLabel: "Calculate Sheet", iconId: "refresh", size: "small" },
          ],
        },
        {
          id: "formulas.solutions",
          label: "Solutions",
          buttons: [
            { id: "formulas.solutions.solver", label: "Solver", ariaLabel: "Solver", iconId: "puzzle", kind: "dropdown", size: "large" },
            { id: "formulas.solutions.analysisToolPak", label: "Analysis ToolPak", ariaLabel: "Analysis ToolPak", iconId: "settings", kind: "dropdown" },
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
            {
              id: "data.getTransform.getData",
              label: "Get Data",
              ariaLabel: "Get Data",
              iconId: "arrowDown",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "data.getTransform.getData.fromFile", label: "From File", ariaLabel: "From File", iconId: "file" },
                { id: "data.getTransform.getData.fromDatabase", label: "From Database", ariaLabel: "From Database", iconId: "folderOpen" },
                { id: "data.getTransform.getData.fromAzure", label: "From Azure", ariaLabel: "From Azure", iconId: "cloud" },
                { id: "data.getTransform.getData.fromOnlineServices", label: "From Online Services", ariaLabel: "From Online Services", iconId: "globe" },
                { id: "data.getTransform.getData.fromOtherSources", label: "From Other Sources", ariaLabel: "From Other Sources", iconId: "plus" },
              ],
            },
            { id: "data.getTransform.recentSources", label: "Recent Sources", ariaLabel: "Recent Sources", iconId: "clock", kind: "dropdown" },
            { id: "data.getTransform.existingConnections", label: "Existing Connections", ariaLabel: "Existing Connections", iconId: "link", kind: "dropdown" },
          ],
        },
        {
          id: "data.queriesConnections",
          label: "Queries & Connections",
          buttons: [
            {
              id: "data.queriesConnections.refreshAll",
              label: "Refresh All",
              ariaLabel: "Refresh All",
              iconId: "refresh",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "data.queriesConnections.refreshAll", label: "Refresh All", ariaLabel: "Refresh All", iconId: "refresh" },
                { id: "data.queriesConnections.refreshAll.refresh", label: "Refresh", ariaLabel: "Refresh", iconId: "refresh" },
                { id: "data.queriesConnections.refreshAll.refreshAllConnections", label: "Refresh All Connections", ariaLabel: "Refresh All Connections", iconId: "link" },
                { id: "data.queriesConnections.refreshAll.refreshAllQueries", label: "Refresh All Queries", ariaLabel: "Refresh All Queries", iconId: "folderOpen" },
              ],
            },
            { id: "data.queriesConnections.queriesConnections", label: "Queries & Connections", ariaLabel: "Queries and Connections", iconId: "folderOpen", kind: "toggle", defaultPressed: false },
            { id: "data.queriesConnections.properties", label: "Properties", ariaLabel: "Properties", iconId: "settings", kind: "dropdown" },
          ],
        },
        {
          id: "data.sortFilter",
          label: "Sort & Filter",
          buttons: [
            { id: "data.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", iconId: "sort" },
            { id: "data.sortFilter.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", iconId: "sort" },
            {
              id: "data.sortFilter.sort",
              label: "Sort",
              ariaLabel: "Sort",
              iconId: "sort",
              kind: "dropdown",
              menuItems: [
                { id: "data.sortFilter.sort.customSort", label: "Custom Sort…", ariaLabel: "Custom Sort", iconId: "settings" },
                { id: "data.sortFilter.sort.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", iconId: "sort" },
                { id: "data.sortFilter.sort.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", iconId: "sort" },
              ],
            },
            { id: "data.sortFilter.filter", label: "Filter", ariaLabel: "Filter", iconId: "filter", kind: "toggle" },
            { id: "data.sortFilter.clear", label: "Clear", ariaLabel: "Clear", iconId: "close" },
            { id: "data.sortFilter.reapply", label: "Reapply", ariaLabel: "Reapply", iconId: "refresh" },
            {
              id: "data.sortFilter.advanced",
              label: "Advanced",
              ariaLabel: "Advanced",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "data.sortFilter.advanced.advancedFilter", label: "Advanced Filter…", ariaLabel: "Advanced Filter", iconId: "settings" },
                { id: "data.sortFilter.advanced.clearFilter", label: "Clear Filter", ariaLabel: "Clear Filter", iconId: "close" },
              ],
            },
          ],
        },
        {
          id: "data.dataTools",
          label: "Data Tools",
          buttons: [
            {
              id: "data.dataTools.textToColumns",
              label: "Text to Columns",
              ariaLabel: "Text to Columns",
              iconId: "insertColumns",
              kind: "dropdown",
              menuItems: [
                { id: "data.dataTools.textToColumns", label: "Text to Columns…", ariaLabel: "Text to Columns", iconId: "insertColumns" },
                { id: "data.dataTools.textToColumns.reapply", label: "Reapply", ariaLabel: "Reapply", iconId: "refresh" },
              ],
            },
            { id: "data.dataTools.flashFill", label: "Flash Fill", ariaLabel: "Flash Fill", iconId: "lightning" },
            {
              id: "data.dataTools.removeDuplicates",
              label: "Remove Duplicates",
              ariaLabel: "Remove Duplicates",
              iconId: "trash",
              kind: "dropdown",
              menuItems: [
                { id: "data.dataTools.removeDuplicates", label: "Remove Duplicates…", ariaLabel: "Remove Duplicates", iconId: "trash" },
                { id: "data.dataTools.removeDuplicates.advanced", label: "Advanced…", ariaLabel: "Advanced", iconId: "settings" },
              ],
            },
            {
              id: "data.dataTools.dataValidation",
              label: "Data Validation",
              ariaLabel: "Data Validation",
              iconId: "check",
              kind: "dropdown",
              menuItems: [
                { id: "data.dataTools.dataValidation", label: "Data Validation…", ariaLabel: "Data Validation", iconId: "check" },
                { id: "data.dataTools.dataValidation.circleInvalid", label: "Circle Invalid Data", ariaLabel: "Circle Invalid Data", iconId: "warning" },
                { id: "data.dataTools.dataValidation.clearCircles", label: "Clear Validation Circles", ariaLabel: "Clear Validation Circles", iconId: "close" },
              ],
            },
            {
              id: "data.dataTools.consolidate",
              label: "Consolidate",
              ariaLabel: "Consolidate",
              iconId: "puzzle",
              kind: "dropdown",
              menuItems: [{ id: "data.dataTools.consolidate", label: "Consolidate…", ariaLabel: "Consolidate", iconId: "puzzle" }],
            },
            {
              id: "data.dataTools.relationships",
              label: "Relationships",
              ariaLabel: "Relationships",
              iconId: "link",
              kind: "dropdown",
              menuItems: [
                { id: "data.dataTools.relationships", label: "Relationships…", ariaLabel: "Relationships", iconId: "link" },
                { id: "data.dataTools.relationships.manage", label: "Manage Relationships…", ariaLabel: "Manage Relationships", iconId: "settings" },
              ],
            },
            {
              id: "data.dataTools.manageDataModel",
              label: "Manage Data Model",
              ariaLabel: "Manage Data Model",
              iconId: "puzzle",
              kind: "dropdown",
              menuItems: [
                { id: "data.dataTools.manageDataModel", label: "Manage Data Model", ariaLabel: "Manage Data Model", iconId: "puzzle" },
                { id: "data.dataTools.manageDataModel.addToDataModel", label: "Add to Data Model", ariaLabel: "Add to Data Model", iconId: "plus" },
              ],
            },
          ],
        },
        {
          id: "data.forecast",
          label: "Forecast",
          buttons: [
            {
              id: "data.forecast.whatIfAnalysis",
              label: "What-If Analysis",
              ariaLabel: "What-If Analysis",
              iconId: "help",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "data.forecast.whatIfAnalysis.scenarioManager", label: "Scenario Manager…", ariaLabel: "Scenario Manager", iconId: "settings" },
                { id: "data.forecast.whatIfAnalysis.goalSeek", label: "Goal Seek…", ariaLabel: "Goal Seek", iconId: "find" },
                { id: "data.forecast.whatIfAnalysis.dataTable", label: "Data Table…", ariaLabel: "Data Table", icon: "▦" },
              ],
            },
            {
              id: "data.forecast.forecastSheet",
              label: "Forecast Sheet",
              ariaLabel: "Forecast Sheet",
              iconId: "chart",
              kind: "dropdown",
              menuItems: [
                { id: "data.forecast.forecastSheet", label: "Forecast Sheet…", ariaLabel: "Forecast Sheet", iconId: "chart" },
                { id: "data.forecast.forecastSheet.options", label: "Options…", ariaLabel: "Forecast Options", iconId: "settings" },
              ],
            },
          ],
        },
        {
          id: "data.outline",
          label: "Outline",
          buttons: [
            {
              id: "data.outline.group",
              label: "Group",
              ariaLabel: "Group",
              iconId: "plus",
              kind: "dropdown",
              menuItems: [
                { id: "data.outline.group.group", label: "Group…", ariaLabel: "Group", iconId: "plus" },
                { id: "data.outline.group.groupSelection", label: "Group Selection", ariaLabel: "Group Selection", icon: "▦" },
              ],
            },
            {
              id: "data.outline.ungroup",
              label: "Ungroup",
              ariaLabel: "Ungroup",
              iconId: "minus",
              kind: "dropdown",
              menuItems: [
                { id: "data.outline.ungroup.ungroup", label: "Ungroup…", ariaLabel: "Ungroup", iconId: "minus" },
                { id: "data.outline.ungroup.clearOutline", label: "Clear Outline", ariaLabel: "Clear Outline", iconId: "close" },
              ],
            },
            {
              id: "data.outline.subtotal",
              label: "Subtotal",
              ariaLabel: "Subtotal",
              iconId: "autoSum",
              kind: "dropdown",
              menuItems: [{ id: "data.outline.subtotal", label: "Subtotal…", ariaLabel: "Subtotal", iconId: "autoSum" }],
            },
            { id: "data.outline.showDetail", label: "Show Detail", ariaLabel: "Show Detail", iconId: "plus" },
            { id: "data.outline.hideDetail", label: "Hide Detail", ariaLabel: "Hide Detail", iconId: "minus" },
          ],
        },
        {
          id: "data.dataTypes",
          label: "Data Types",
          buttons: [
            { id: "data.dataTypes.stocks", label: "Stocks", ariaLabel: "Stocks", iconId: "chart", kind: "dropdown", size: "large" },
            { id: "data.dataTypes.geography", label: "Geography", ariaLabel: "Geography", iconId: "globe", kind: "dropdown", size: "large" },
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
            {
              id: "review.proofing.spelling",
              label: "Spelling",
              ariaLabel: "Spelling",
              iconId: "check",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "review.proofing.spelling", label: "Spelling", ariaLabel: "Spelling", iconId: "check" },
                { id: "review.proofing.spelling.thesaurus", label: "Thesaurus", ariaLabel: "Thesaurus", iconId: "help" },
                { id: "review.proofing.spelling.wordCount", label: "Word Count", ariaLabel: "Word Count", iconId: "hash" },
              ],
            },
            { id: "review.proofing.accessibility", label: "Check Accessibility", ariaLabel: "Check Accessibility", iconId: "help", kind: "dropdown" },
            { id: "review.proofing.smartLookup", label: "Smart Lookup", ariaLabel: "Smart Lookup", iconId: "search", kind: "dropdown" },
          ],
        },
        {
          id: "review.comments",
          label: "Comments",
          buttons: [
            { id: "review.comments.newComment", label: "New Comment", ariaLabel: "New Comment", iconId: "comment", size: "large" },
            {
              id: "review.comments.deleteComment",
              label: "Delete",
              ariaLabel: "Delete Comment",
              iconId: "trash",
              kind: "dropdown",
              menuItems: [
                { id: "review.comments.deleteComment", label: "Delete Comment", ariaLabel: "Delete Comment", iconId: "trash" },
                { id: "review.comments.deleteComment.deleteThread", label: "Delete Thread", ariaLabel: "Delete Thread", iconId: "comment" },
                { id: "review.comments.deleteComment.deleteAll", label: "Delete All Comments", ariaLabel: "Delete All Comments", iconId: "trash" },
              ],
            },
            { id: "review.comments.previous", label: "Previous", ariaLabel: "Previous Comment", iconId: "arrowUp" },
            { id: "review.comments.next", label: "Next", ariaLabel: "Next Comment", iconId: "arrowDown" },
            { id: "review.comments.showComments", label: "Show Comments", ariaLabel: "Show Comments", iconId: "eye", kind: "toggle" },
          ],
        },
        {
          id: "review.notes",
          label: "Notes",
          buttons: [
            {
              id: "review.notes.newNote",
              label: "New Note",
              ariaLabel: "New Note",
              iconId: "file",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "review.notes.newNote", label: "New Note", ariaLabel: "New Note", iconId: "file" },
                { id: "review.notes.editNote", label: "Edit Note", ariaLabel: "Edit Note", iconId: "edit" },
              ],
            },
            { id: "review.notes.showAllNotes", label: "Show All Notes", ariaLabel: "Show All Notes", iconId: "eye", kind: "toggle" },
            { id: "review.notes.showHideNote", label: "Show/Hide Note", ariaLabel: "Show or Hide Note", iconId: "eyeOff", kind: "toggle" },
          ],
        },
        {
          id: "review.protect",
          label: "Protect",
          buttons: [
            {
              id: "review.protect.protectSheet",
              label: "Protect Sheet",
              ariaLabel: "Protect Sheet",
              iconId: "lock",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "review.protect.protectSheet", label: "Protect Sheet…", ariaLabel: "Protect Sheet", iconId: "lock" },
                { id: "review.protect.unprotectSheet", label: "Unprotect Sheet…", ariaLabel: "Unprotect Sheet", iconId: "unlock" },
              ],
            },
            {
              id: "review.protect.protectWorkbook",
              label: "Protect Workbook",
              ariaLabel: "Protect Workbook",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "review.protect.protectWorkbook", label: "Protect Workbook…", ariaLabel: "Protect Workbook", iconId: "settings" },
                { id: "review.protect.unprotectWorkbook", label: "Unprotect Workbook…", ariaLabel: "Unprotect Workbook", iconId: "unlock" },
              ],
            },
            {
              id: "review.protect.allowEditRanges",
              label: "Allow Edit Ranges",
              ariaLabel: "Allow Edit Ranges",
              iconId: "check",
              kind: "dropdown",
              menuItems: [
                { id: "review.protect.allowEditRanges", label: "Allow Users to Edit Ranges…", ariaLabel: "Allow Users to Edit Ranges", iconId: "check" },
                { id: "review.protect.allowEditRanges.new", label: "New…", ariaLabel: "New allowed range", iconId: "plus" },
              ],
            },
          ],
        },
        {
          id: "review.ink",
          label: "Ink",
          buttons: [
            { id: "review.ink.startInking", label: "Start Inking", ariaLabel: "Start Inking", iconId: "edit", kind: "toggle", size: "large" },
          ],
        },
        {
          id: "review.language",
          label: "Language",
          buttons: [
            {
              id: "review.language.translate",
              label: "Translate",
              ariaLabel: "Translate",
              iconId: "globe",
              kind: "dropdown",
              menuItems: [
                { id: "review.language.translate.translateSelection", label: "Translate Selection", ariaLabel: "Translate Selection", iconId: "globe" },
                { id: "review.language.translate.translateSheet", label: "Translate Sheet", ariaLabel: "Translate Sheet", iconId: "file" },
              ],
            },
            {
              id: "review.language.language",
              label: "Language",
              ariaLabel: "Language",
              iconId: "globe",
              kind: "dropdown",
              menuItems: [
                { id: "review.language.language.setProofing", label: "Set Proofing Language…", ariaLabel: "Set Proofing Language", iconId: "globe" },
                { id: "review.language.language.translate", label: "Translate", ariaLabel: "Translate", iconId: "globe" },
              ],
            },
          ],
        },
        {
          id: "review.changes",
          label: "Changes",
          buttons: [
            {
              id: "review.changes.trackChanges",
              label: "Track Changes",
              ariaLabel: "Track Changes",
              iconId: "edit",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "review.changes.trackChanges", label: "Track Changes…", ariaLabel: "Track Changes", iconId: "edit" },
                { id: "review.changes.trackChanges.highlight", label: "Highlight Changes…", ariaLabel: "Highlight Changes", iconId: "fillColor" },
              ],
            },
            {
              id: "review.changes.shareWorkbook",
              label: "Share Workbook",
              ariaLabel: "Share Workbook",
              iconId: "users",
              kind: "dropdown",
              menuItems: [
                { id: "review.changes.shareWorkbook", label: "Share Workbook…", ariaLabel: "Share Workbook", iconId: "users" },
                { id: "review.changes.shareWorkbook.shareNow", label: "Share Now", ariaLabel: "Share Now", iconId: "link" },
              ],
            },
            {
              id: "review.changes.protectShareWorkbook",
              label: "Protect and Share Workbook",
              ariaLabel: "Protect and Share Workbook",
              iconId: "lock",
              kind: "dropdown",
              menuItems: [
                { id: "review.changes.protectShareWorkbook", label: "Protect and Share Workbook…", ariaLabel: "Protect and Share Workbook", iconId: "lock" },
                { id: "review.changes.protectShareWorkbook.protectWorkbook", label: "Protect Workbook", ariaLabel: "Protect Workbook", iconId: "settings" },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "view",
      label: "View",
      groups: [
        {
          id: "view.panels",
          label: "Panels",
          buttons: [
            {
              id: "open-marketplace-panel",
              label: "Marketplace",
              ariaLabel: "Marketplace",
              iconId: "puzzle",
              testId: "open-marketplace-panel",
            },
            {
              id: "open-version-history-panel",
              label: "Version History",
              ariaLabel: "Toggle Version History panel",
              iconId: "clock",
              testId: "open-version-history-panel",
            },
            {
              id: "open-branch-manager-panel",
              label: "Branches",
              ariaLabel: "Toggle Branch Manager panel",
              iconId: "shuffle",
              testId: "open-branch-manager-panel",
            },
          ],
        },
        {
          id: "view.appearance",
          label: "Appearance",
          buttons: [
            {
              id: "view.appearance.theme",
              label: "Theme",
              ariaLabel: "Theme",
              iconId: "palette",
              kind: "dropdown",
              testId: "theme-selector",
              menuItems: [
                {
                  id: "view.appearance.theme.system",
                  label: "System",
                  ariaLabel: "Use system theme",
                  iconId: "window",
                  testId: "theme-option-system",
                },
                {
                  id: "view.appearance.theme.light",
                  label: "Light",
                  ariaLabel: "Use light theme",
                  iconId: "chart",
                  testId: "theme-option-light",
                },
                {
                  id: "view.appearance.theme.dark",
                  label: "Dark",
                  ariaLabel: "Use dark theme",
                  iconId: "palette",
                  testId: "theme-option-dark",
                },
                {
                  id: "view.appearance.theme.highContrast",
                  label: "High Contrast",
                  ariaLabel: "Use high contrast theme",
                  icon: "◧",
                  testId: "theme-option-high-contrast",
                },
              ],
            },
          ],
        },
        {
          id: "view.show",
          label: "Show",
          buttons: [
            { id: "view.show.ruler", label: "Ruler", ariaLabel: "Ruler", iconId: "ruler", kind: "toggle", defaultPressed: false },
            { id: "view.show.gridlines", label: "Gridlines", ariaLabel: "Gridlines", icon: "▦", kind: "toggle", defaultPressed: true },
            { id: "view.show.formulaBar", label: "Formula Bar", ariaLabel: "Formula Bar", iconId: "function", kind: "toggle", defaultPressed: true },
            { id: "view.show.headings", label: "Headings", ariaLabel: "Headings", icon: "A1", kind: "toggle", defaultPressed: true },
            {
              id: "view.show.showFormulas",
              label: "Show Formulas",
              ariaLabel: "Show formulas (Ctrl/Cmd+`)",
              icon: "`",
              kind: "toggle",
              defaultPressed: false,
              testId: "ribbon-show-formulas",
            },
            {
              id: "view.show.performanceStats",
              label: "Performance Stats",
              ariaLabel: "Performance stats",
              iconId: "chart",
              kind: "toggle",
              defaultPressed: false,
              testId: "ribbon-perf-stats",
            },
          ],
        },
        {
          id: "view.zoom",
          label: "Zoom",
          buttons: [
            {
              id: "view.zoom.zoom",
              label: "Zoom",
              ariaLabel: "Zoom",
              iconId: "search",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "view.zoom.zoom.400", label: "400%", ariaLabel: "Zoom 400%", icon: "400%" },
                { id: "view.zoom.zoom.200", label: "200%", ariaLabel: "Zoom 200%", icon: "200%" },
                { id: "view.zoom.zoom.150", label: "150%", ariaLabel: "Zoom 150%", icon: "150%" },
                { id: "view.zoom.zoom.100", label: "100%", ariaLabel: "Zoom 100%", icon: "100%" },
                { id: "view.zoom.zoom.75", label: "75%", ariaLabel: "Zoom 75%", icon: "75%" },
                { id: "view.zoom.zoom.50", label: "50%", ariaLabel: "Zoom 50%", icon: "50%" },
                { id: "view.zoom.zoom.25", label: "25%", ariaLabel: "Zoom 25%", icon: "25%" },
                { id: "view.zoom.zoom.custom", label: "Custom…", ariaLabel: "Custom Zoom", iconId: "settings" },
              ],
            },
            { id: "view.zoom.zoom100", label: "100%", ariaLabel: "Zoom to 100%", icon: "100%" },
            { id: "view.zoom.zoomToSelection", label: "Zoom to Selection", ariaLabel: "Zoom to Selection", iconId: "find" },
          ],
        },
        {
          id: "view.workbookViews",
          label: "Workbook Views",
          buttons: [
            { id: "view.workbookViews.normal", label: "Normal", ariaLabel: "Normal View", icon: "▦", kind: "toggle", defaultPressed: true, size: "large" },
            { id: "view.workbookViews.pageBreakPreview", label: "Page Break Preview", ariaLabel: "Page Break Preview", icon: "⤶", kind: "toggle", size: "large" },
            { id: "view.workbookViews.pageLayout", label: "Page Layout", ariaLabel: "Page Layout View", iconId: "file", kind: "toggle", size: "large" },
            {
              id: "view.workbookViews.customViews",
              label: "Custom Views",
              ariaLabel: "Custom Views",
              iconId: "eye",
              kind: "dropdown",
              menuItems: [
                { id: "view.workbookViews.customViews", label: "Custom Views…", ariaLabel: "Custom Views", iconId: "eye" },
                { id: "view.workbookViews.customViews.manage", label: "Manage Views…", ariaLabel: "Manage Views", iconId: "settings" },
              ],
            },
          ],
        },
        {
          id: "view.window",
          label: "Window",
          buttons: [
            {
              id: "view.window.newWindow",
              label: "New Window",
              ariaLabel: "New Window",
              iconId: "window",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "view.window.newWindow", label: "New Window", ariaLabel: "New Window", iconId: "window" },
                { id: "view.window.newWindow.newWindowForActiveSheet", label: "New Window for Active Sheet", ariaLabel: "New Window for Active Sheet", iconId: "file" },
              ],
            },
            {
              id: "view.window.arrangeAll",
              label: "Arrange All",
              ariaLabel: "Arrange All",
              iconId: "window",
              kind: "dropdown",
              menuItems: [
                { id: "view.window.arrangeAll", label: "Arrange All…", ariaLabel: "Arrange All", iconId: "window" },
                { id: "view.window.arrangeAll.tiled", label: "Tiled", ariaLabel: "Tiled", icon: "▦" },
                { id: "view.window.arrangeAll.horizontal", label: "Horizontal", ariaLabel: "Horizontal", iconId: "arrowLeftRight" },
                { id: "view.window.arrangeAll.vertical", label: "Vertical", ariaLabel: "Vertical", iconId: "arrowUpDown" },
                { id: "view.window.arrangeAll.cascade", label: "Cascade", ariaLabel: "Cascade", iconId: "window" },
              ],
            },
            {
              id: "view.window.freezePanes",
              label: "Freeze Panes",
              ariaLabel: "Freeze Panes",
              iconId: "lock",
              kind: "dropdown",
              menuItems: [
                { id: "view.window.freezePanes.freezePanes", label: "Freeze Panes", ariaLabel: "Freeze Panes", iconId: "lock" },
                { id: "view.window.freezePanes.freezeTopRow", label: "Freeze Top Row", ariaLabel: "Freeze Top Row", iconId: "arrowUp" },
                { id: "view.window.freezePanes.freezeFirstColumn", label: "Freeze First Column", ariaLabel: "Freeze First Column", iconId: "arrowLeft" },
                { id: "view.window.freezePanes.unfreeze", label: "Unfreeze Panes", ariaLabel: "Unfreeze Panes", iconId: "close" },
              ],
            },
            { id: "view.window.split", label: "Split", ariaLabel: "Split", iconId: "divide", kind: "toggle" },
            { id: "view.window.hide", label: "Hide", ariaLabel: "Hide", iconId: "eyeOff" },
            { id: "view.window.unhide", label: "Unhide", ariaLabel: "Unhide", iconId: "eye" },
            { id: "view.window.viewSideBySide", label: "View Side by Side", ariaLabel: "View Side by Side", iconId: "sideBySide", kind: "toggle" },
            { id: "view.window.synchronousScrolling", label: "Synchronous Scrolling", ariaLabel: "Synchronous Scrolling", iconId: "syncScroll", kind: "toggle" },
            { id: "view.window.resetWindowPosition", label: "Reset Window Position", ariaLabel: "Reset Window Position", iconId: "undo" },
            {
              id: "view.window.switchWindows",
              label: "Switch Windows",
              ariaLabel: "Switch Windows",
              iconId: "refresh",
              kind: "dropdown",
              menuItems: [
                { id: "view.window.switchWindows", label: "Switch Windows…", ariaLabel: "Switch Windows", iconId: "refresh" },
                { id: "view.window.switchWindows.window1", label: "Book1.xlsx", ariaLabel: "Switch to Book1", iconId: "file" },
                { id: "view.window.switchWindows.window2", label: "Forecast.xlsx", ariaLabel: "Switch to Forecast", iconId: "file" },
              ],
            },
          ],
        },
        {
          id: "view.macros",
          label: "Macros",
          buttons: [
            {
              id: "view.macros.viewMacros",
              label: "View Macros",
              ariaLabel: "View Macros",
              iconId: "file",
              kind: "dropdown",
              size: "large",
              testId: "ribbon-view-macros",
              menuItems: [
                { id: "view.macros.viewMacros", label: "View Macros…", ariaLabel: "View Macros", iconId: "file", testId: "ribbon-view-macros-open" },
                { id: "view.macros.viewMacros.run", label: "Run…", ariaLabel: "Run Macro", iconId: "play", testId: "ribbon-view-macros-run" },
                { id: "view.macros.viewMacros.edit", label: "Edit…", ariaLabel: "Edit Macro", iconId: "edit", testId: "ribbon-view-macros-edit" },
                { id: "view.macros.viewMacros.delete", label: "Delete…", ariaLabel: "Delete Macro", iconId: "trash", testId: "ribbon-view-macros-delete" },
              ],
            },
            {
              id: "view.macros.recordMacro",
              label: "Record Macro",
              ariaLabel: "Record Macro",
              iconId: "record",
              kind: "dropdown",
              testId: "ribbon-view-record-macro",
              menuItems: [
                { id: "view.macros.recordMacro", label: "Record Macro…", ariaLabel: "Record Macro", iconId: "record", testId: "ribbon-view-record-macro-start" },
                { id: "view.macros.recordMacro.stop", label: "Stop Recording", ariaLabel: "Stop Recording", iconId: "stop", testId: "ribbon-view-record-macro-stop" },
              ],
            },
            { id: "view.macros.useRelativeReferences", label: "Use Relative References", ariaLabel: "Use Relative References", iconId: "pin", kind: "toggle" },
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
            {
              id: "developer.code.visualBasic",
              label: "Visual Basic",
              ariaLabel: "Visual Basic",
              icon: "VB",
              size: "large",
              testId: "ribbon-developer-visual-basic",
            },
            {
              id: "developer.code.macros",
              label: "Macros",
              ariaLabel: "Macros",
              iconId: "file",
              kind: "dropdown",
              size: "large",
              testId: "ribbon-developer-macros",
              menuItems: [
                { id: "developer.code.macros", label: "Macros…", ariaLabel: "Macros", iconId: "file", testId: "ribbon-developer-macros-open" },
                { id: "developer.code.macros.run", label: "Run…", ariaLabel: "Run Macro", iconId: "play", testId: "ribbon-developer-macros-run" },
                { id: "developer.code.macros.edit", label: "Edit…", ariaLabel: "Edit Macro", iconId: "edit", testId: "ribbon-developer-macros-edit" },
              ],
            },
            {
              id: "developer.code.recordMacro",
              label: "Record Macro",
              ariaLabel: "Record Macro",
              iconId: "record",
              kind: "dropdown",
              testId: "ribbon-developer-record-macro",
              menuItems: [
                { id: "developer.code.recordMacro", label: "Record Macro…", ariaLabel: "Record Macro", iconId: "record", testId: "ribbon-developer-record-macro-start" },
                { id: "developer.code.recordMacro.stop", label: "Stop Recording", ariaLabel: "Stop Recording", iconId: "stop", testId: "ribbon-developer-record-macro-stop" },
              ],
            },
            { id: "developer.code.useRelativeReferences", label: "Use Relative References", ariaLabel: "Use Relative References", iconId: "pin", kind: "toggle" },
            {
              id: "developer.code.macroSecurity",
              label: "Macro Security",
              ariaLabel: "Macro Security",
              iconId: "lock",
              kind: "dropdown",
              testId: "ribbon-developer-macro-security",
              menuItems: [
                { id: "developer.code.macroSecurity", label: "Macro Security…", ariaLabel: "Macro Security", iconId: "lock", testId: "ribbon-developer-macro-security-open" },
                { id: "developer.code.macroSecurity.trustCenter", label: "Trust Center…", ariaLabel: "Trust Center", iconId: "lock", testId: "ribbon-developer-macro-security-trust-center" },
              ],
            },
          ],
        },
        {
          id: "developer.addins",
          label: "Add-ins",
          buttons: [
            {
              id: "developer.addins.addins",
              label: "Add-ins",
              ariaLabel: "Add-ins",
              iconId: "puzzle",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "developer.addins.addins.excelAddins", label: "Excel Add-ins…", ariaLabel: "Excel Add-ins", iconId: "puzzle" },
                { id: "developer.addins.addins.browse", label: "Browse…", ariaLabel: "Browse Add-ins", iconId: "folderOpen" },
                { id: "developer.addins.addins.manage", label: "Manage…", ariaLabel: "Manage Add-ins", iconId: "settings" },
              ],
            },
            {
              id: "developer.addins.comAddins",
              label: "COM Add-ins",
              ariaLabel: "COM Add-ins",
              iconId: "puzzle",
              kind: "dropdown",
              menuItems: [{ id: "developer.addins.comAddins", label: "COM Add-ins…", ariaLabel: "COM Add-ins", iconId: "puzzle" }],
            },
          ],
        },
        {
          id: "developer.controls",
          label: "Controls",
          buttons: [
            {
              id: "developer.controls.insert",
              label: "Insert",
              ariaLabel: "Insert Control",
              iconId: "plus",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "developer.controls.insert.button", label: "Button", ariaLabel: "Insert Button", iconId: "check" },
                { id: "developer.controls.insert.checkbox", label: "Check Box", ariaLabel: "Insert Check Box", iconId: "check" },
                { id: "developer.controls.insert.combobox", label: "Combo Box", ariaLabel: "Insert Combo Box", iconId: "arrowDown" },
                { id: "developer.controls.insert.listbox", label: "List Box", ariaLabel: "Insert List Box", iconId: "menu" },
                { id: "developer.controls.insert.scrollbar", label: "Scroll Bar", ariaLabel: "Insert Scroll Bar", iconId: "arrowUpDown" },
                { id: "developer.controls.insert.spinButton", label: "Spin Button", ariaLabel: "Insert Spin Button", iconId: "arrowUpDown" },
              ],
            },
            { id: "developer.controls.designMode", label: "Design Mode", ariaLabel: "Design Mode", iconId: "settings", kind: "toggle" },
            {
              id: "developer.controls.properties",
              label: "Properties",
              ariaLabel: "Properties",
              iconId: "settings",
              kind: "dropdown",
              menuItems: [
                { id: "developer.controls.properties", label: "Properties…", ariaLabel: "Properties", iconId: "settings" },
                { id: "developer.controls.properties.viewProperties", label: "View Properties", ariaLabel: "View Properties", iconId: "eye" },
              ],
            },
            { id: "developer.controls.viewCode", label: "View Code", ariaLabel: "View Code", iconId: "code" },
            { id: "developer.controls.runDialog", label: "Run Dialog", ariaLabel: "Run Dialog", iconId: "play" },
          ],
        },
        {
          id: "developer.xml",
          label: "XML",
          buttons: [
            {
              id: "developer.xml.source",
              label: "Source",
              ariaLabel: "XML Source",
              iconId: "code",
              kind: "dropdown",
              size: "large",
              menuItems: [
                { id: "developer.xml.source", label: "XML Source", ariaLabel: "XML Source", iconId: "code" },
                { id: "developer.xml.source.refresh", label: "Refresh XML Data", ariaLabel: "Refresh XML Data", iconId: "refresh" },
              ],
            },
            {
              id: "developer.xml.mapProperties",
              label: "Map Properties",
              ariaLabel: "Map Properties",
              iconId: "globe",
              kind: "dropdown",
              menuItems: [{ id: "developer.xml.mapProperties", label: "Map Properties…", ariaLabel: "Map Properties", iconId: "globe" }],
            },
            { id: "developer.xml.import", label: "Import", ariaLabel: "Import XML", iconId: "arrowDown" },
            { id: "developer.xml.export", label: "Export", ariaLabel: "Export XML", iconId: "arrowUp" },
            { id: "developer.xml.refreshData", label: "Refresh Data", ariaLabel: "Refresh Data", iconId: "refresh" },
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
            { id: "help.support.help", label: "Help", ariaLabel: "Help", iconId: "help", kind: "dropdown", size: "large" },
            { id: "help.support.training", label: "Training", ariaLabel: "Training", iconId: "help", kind: "dropdown" },
            { id: "help.support.contactSupport", label: "Contact Support", ariaLabel: "Contact Support", iconId: "help", kind: "dropdown" },
            { id: "help.support.feedback", label: "Feedback", ariaLabel: "Feedback", iconId: "edit", kind: "dropdown" },
          ],
        },
      ],
    },
  ],
};
