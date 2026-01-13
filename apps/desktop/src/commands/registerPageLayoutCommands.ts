import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { PageSetup } from "../print/index.js";

export const PAGE_LAYOUT_COMMANDS = {
  pageSetupDialog: "pageLayout.pageSetup.pageSetupDialog",
  margins: {
    normal: "pageLayout.pageSetup.margins.normal",
    wide: "pageLayout.pageSetup.margins.wide",
    narrow: "pageLayout.pageSetup.margins.narrow",
    custom: "pageLayout.pageSetup.margins.custom",
  },
  orientation: {
    portrait: "pageLayout.pageSetup.orientation.portrait",
    landscape: "pageLayout.pageSetup.orientation.landscape",
  },
  size: {
    letter: "pageLayout.pageSetup.size.letter",
    a4: "pageLayout.pageSetup.size.a4",
    more: "pageLayout.pageSetup.size.more",
  },
  printArea: {
    setPrintArea: "pageLayout.printArea.setPrintArea",
    clearPrintArea: "pageLayout.printArea.clearPrintArea",
    set: "pageLayout.pageSetup.printArea.set",
    clear: "pageLayout.pageSetup.printArea.clear",
    addTo: "pageLayout.pageSetup.printArea.addTo",
  },
  exportPdf: "pageLayout.export.exportPdf",
} as const;

export type PageLayoutCommandHandlers = {
  openPageSetupDialog: () => void | Promise<void>;
  updatePageSetup: (patch: (current: PageSetup) => PageSetup) => void | Promise<void>;
  setPrintArea: () => void | Promise<void>;
  clearPrintArea: () => void | Promise<void>;
  addToPrintArea: () => void | Promise<void>;
  exportPdf: () => void | Promise<void>;
};

export function registerPageLayoutCommands(params: { commandRegistry: CommandRegistry; handlers: PageLayoutCommandHandlers }): void {
  const { commandRegistry, handlers } = params;

  const category = "Page Layout";

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.pageSetupDialog, "Page Setup…", () => handlers.openPageSetupDialog(), {
    category,
    icon: null,
    description: "Open the Page Setup dialog",
    keywords: ["page setup", "print", "margins", "orientation", "paper size"],
  });

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.margins.normal,
    "Margins: Normal",
    () =>
      handlers.updatePageSetup((current) => ({
        ...current,
        margins: { ...current.margins, left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
      })),
    {
      category,
      icon: null,
      description: "Set page margins to the Normal preset",
      keywords: ["margins", "page setup", "normal"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.margins.wide,
    "Margins: Wide",
    () =>
      handlers.updatePageSetup((current) => ({
        ...current,
        margins: { ...current.margins, left: 1, right: 1, top: 1, bottom: 1 },
      })),
    {
      category,
      icon: null,
      description: "Set page margins to the Wide preset",
      keywords: ["margins", "page setup", "wide"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.margins.narrow,
    "Margins: Narrow",
    () =>
      handlers.updatePageSetup((current) => ({
        ...current,
        margins: { ...current.margins, left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 },
      })),
    {
      category,
      icon: null,
      description: "Set page margins to the Narrow preset",
      keywords: ["margins", "page setup", "narrow"],
    },
  );

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.margins.custom, "Margins: Custom…", () => handlers.openPageSetupDialog(), {
    category,
    icon: null,
    description: "Open the Page Setup dialog to edit margins",
    keywords: ["margins", "page setup", "custom"],
  });

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.orientation.portrait,
    "Orientation: Portrait",
    () => handlers.updatePageSetup((current) => ({ ...current, orientation: "portrait" })),
    {
      category,
      icon: null,
      description: "Set page orientation to Portrait",
      keywords: ["page setup", "orientation", "portrait"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.orientation.landscape,
    "Orientation: Landscape",
    () => handlers.updatePageSetup((current) => ({ ...current, orientation: "landscape" })),
    {
      category,
      icon: null,
      description: "Set page orientation to Landscape",
      keywords: ["page setup", "orientation", "landscape"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.size.letter,
    "Paper Size: Letter",
    () => handlers.updatePageSetup((current) => ({ ...current, paperSize: 1 })),
    {
      category,
      icon: null,
      description: "Set paper size to Letter",
      keywords: ["page setup", "paper size", "letter"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.size.a4,
    "Paper Size: A4",
    () => handlers.updatePageSetup((current) => ({ ...current, paperSize: 9 })),
    {
      category,
      icon: null,
      description: "Set paper size to A4",
      keywords: ["page setup", "paper size", "a4"],
    },
  );

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.size.more, "Paper Size: More…", () => handlers.openPageSetupDialog(), {
    category,
    icon: null,
    description: "Open the Page Setup dialog to change paper size",
    keywords: ["page setup", "paper size", "more"],
  });

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.printArea.setPrintArea, "Set Print Area", () => handlers.setPrintArea(), {
    category,
    icon: null,
    description: "Set the print area to the current selection",
    keywords: ["print area", "page layout", "print"],
  });

  // Ribbon schema currently includes a second "Set Print Area" id under the Page Setup dropdown.
  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.printArea.set, "Set Print Area", () => handlers.setPrintArea(), {
    category,
    icon: null,
    description: "Set the print area to the current selection",
    keywords: ["print area", "page setup", "print"],
  });

  commandRegistry.registerBuiltinCommand(
    PAGE_LAYOUT_COMMANDS.printArea.clearPrintArea,
    "Clear Print Area",
    () => handlers.clearPrintArea(),
    {
      category,
      icon: null,
      description: "Clear the sheet print area",
      keywords: ["print area", "clear", "page layout", "print"],
    },
  );

  // Ribbon schema currently includes a second "Clear Print Area" id under the Page Setup dropdown.
  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.printArea.clear, "Clear Print Area", () => handlers.clearPrintArea(), {
    category,
    icon: null,
    description: "Clear the sheet print area",
    keywords: ["print area", "clear", "page setup", "print"],
  });

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.printArea.addTo, "Add to Print Area", () => handlers.addToPrintArea(), {
    category,
    icon: null,
    description: "Add the current selection to the sheet print area",
    keywords: ["print area", "add", "page layout", "print"],
  });

  commandRegistry.registerBuiltinCommand(PAGE_LAYOUT_COMMANDS.exportPdf, "Export to PDF", () => handlers.exportPdf(), {
    category,
    icon: null,
    description: "Export the current sheet to a PDF file",
    keywords: ["export", "pdf", "print"],
  });
}

