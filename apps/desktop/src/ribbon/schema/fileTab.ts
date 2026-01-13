import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const fileTab: RibbonTabDefinition = {
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
            {
              id: "file.info.manageWorkbook.recoverUnsaved",
              label: "Recover Unsaved Workbooks…",
              ariaLabel: "Recover Unsaved Workbooks",
              iconId: "clock",
            },
            { id: "file.info.manageWorkbook.versions", label: "Version History", ariaLabel: "Version History", iconId: "clock" },
            { id: "file.info.manageWorkbook.branches", label: "Branches", ariaLabel: "Branches", iconId: "shuffle" },
            {
              id: "file.info.manageWorkbook.properties",
              label: "Properties",
              ariaLabel: "Properties",
              iconId: "settings",
            },
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
            { id: "file.save.saveAs.download", label: "Download a Copy", ariaLabel: "Download a Copy", iconId: "download" },
          ],
        },
        { id: "file.save.autoSave", label: "AutoSave", ariaLabel: "AutoSave", iconId: "cloud", kind: "toggle", defaultPressed: false },
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
};
