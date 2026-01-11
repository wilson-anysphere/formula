export type CommandContribution = {
  commandId: string;
  title: string;
  category: string | null;
  icon?: string | null;
  source: { kind: "builtin" } | { kind: "extension"; extensionId: string };
};

export class CommandRegistry {
  private readonly commands = new Map<string, CommandContribution & { run: (...args: any[]) => Promise<any> }>();
  private readonly listeners = new Set<() => void>();

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of [...this.listeners]) {
      try {
        listener();
      } catch {
        // ignore
      }
    }
  }

  registerBuiltinCommand(commandId: string, title: string, run: (...args: any[]) => any): void {
    const id = String(commandId);
    this.commands.set(id, {
      commandId: id,
      title,
      category: null,
      icon: null,
      source: { kind: "builtin" },
      run: async (...args) => run(...args),
    });
    this.emit();
  }

  /**
   * Replace all extension commands in the registry with the given set.
   */
  setExtensionCommands(
    commands: Array<{ extensionId: string; command: string; title: string; category?: string | null; icon?: string | null }>,
    executor: (commandId: string, ...args: any[]) => Promise<any>,
  ): void {
    // Remove existing extension commands (keep builtin).
    for (const [id, entry] of this.commands.entries()) {
      if (entry.source.kind === "extension") this.commands.delete(id);
    }

    for (const cmd of commands) {
      const id = String(cmd.command);
      if (this.commands.has(id)) {
        // Deterministic failure: keep the first registered command (builtin).
        throw new Error(`Duplicate command id: ${id}`);
      }
      this.commands.set(id, {
        commandId: id,
        title: String(cmd.title ?? id),
        category: cmd.category ?? null,
        icon: cmd.icon ?? null,
        source: { kind: "extension", extensionId: String(cmd.extensionId) },
        run: async (...args) => executor(id, ...args),
      });
    }

    this.emit();
  }

  listCommands(): CommandContribution[] {
    return [...this.commands.values()].map(({ run: _run, ...rest }) => rest);
  }

  async executeCommand(commandId: string, ...args: any[]): Promise<any> {
    const entry = this.commands.get(String(commandId));
    if (!entry) throw new Error(`Unknown command: ${commandId}`);
    return entry.run(...args);
  }
}

