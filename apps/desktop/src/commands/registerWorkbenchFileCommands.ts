import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";

export const WORKBENCH_FILE_COMMANDS = {
  newWorkbook: "workbench.newWorkbook",
  openWorkbook: "workbench.openWorkbook",
  saveWorkbook: "workbench.saveWorkbook",
  saveWorkbookAs: "workbench.saveWorkbookAs",
  print: "workbench.print",
  printPreview: "workbench.printPreview",
  closeWorkbook: "workbench.closeWorkbook",
  quit: "workbench.quit",
} as const;

export type WorkbenchFileCommandHandlers = {
  newWorkbook: () => void | Promise<void>;
  openWorkbook: () => void | Promise<void>;
  saveWorkbook: () => void | Promise<void>;
  saveWorkbookAs: () => void | Promise<void>;
  print: () => void | Promise<void>;
  printPreview: () => void | Promise<void>;
  closeWorkbook: () => void | Promise<void>;
  quit: () => void | Promise<void>;
};

export function registerWorkbenchFileCommands(params: {
  commandRegistry: CommandRegistry;
  handlers: WorkbenchFileCommandHandlers;
}): void {
  const { commandRegistry, handlers } = params;

  const category = t("menu.file");

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.newWorkbook,
    t("command.workbench.newWorkbook"),
    () => handlers.newWorkbook(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.newWorkbook"),
      keywords: ["new", "create", "workbook"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.openWorkbook,
    t("command.workbench.openWorkbook"),
    () => handlers.openWorkbook(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.openWorkbook"),
      keywords: ["open", "workbook", "file"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.saveWorkbook,
    t("command.workbench.saveWorkbook"),
    () => handlers.saveWorkbook(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.saveWorkbook"),
      keywords: ["save", "workbook", "file"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.saveWorkbookAs,
    t("command.workbench.saveWorkbookAs"),
    () => handlers.saveWorkbookAs(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.saveWorkbookAs"),
      keywords: ["save", "save as", "workbook", "file"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.print,
    t("command.workbench.print"),
    () => handlers.print(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.print"),
      keywords: ["print", "pdf", "page setup"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.printPreview,
    t("command.workbench.printPreview"),
    () => handlers.printPreview(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.printPreview"),
      keywords: ["print preview", "preview", "pdf", "page setup"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.closeWorkbook,
    t("command.workbench.closeWorkbook"),
    () => handlers.closeWorkbook(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.closeWorkbook"),
      keywords: ["close", "workbook", "window"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    WORKBENCH_FILE_COMMANDS.quit,
    t("command.workbench.quit"),
    () => handlers.quit(),
    {
      category,
      icon: null,
      description: t("commandDescription.workbench.quit"),
      keywords: ["quit", "exit", "close"],
    },
  );
}
