export type CommandContribution = {
  commandId: string;
  title: string;
  category: string | null;
  icon?: string | null;
  description?: string | null;
  keywords?: string[] | null;
  /**
   * Optional context key expression that determines whether this command should be
   * visible in context-aware UI surfaces (e.g. the command palette).
   *
   * Note: `CommandRegistry` itself does not evaluate this expression when executing
   * commands; callers should enforce permissions at the command implementation
   * layer as needed.
   */
  when?: string | null;
  source: { kind: "builtin" } | { kind: "extension"; extensionId: string };
};

export type CommandExecutionEvent = {
  commandId: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  args: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  result?: any;
  error?: unknown;
};

export type CommandWillExecuteEvent = {
  commandId: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  args: any[];
  command: CommandContribution;
};

export class CommandRegistry {
  private readonly commands = new Map<string, CommandContribution & { run: (...args: any[]) => Promise<any> }>();
  private readonly listeners = new Set<() => void>();
  private readonly willExecuteListeners = new Set<(evt: CommandWillExecuteEvent) => void | Promise<void>>();
  private readonly executeListeners = new Set<(evt: CommandExecutionEvent) => void>();

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  onDidExecuteCommand(listener: (evt: CommandExecutionEvent) => void): () => void {
    this.executeListeners.add(listener);
    return () => this.executeListeners.delete(listener);
  }

  onWillExecuteCommand(listener: (evt: CommandWillExecuteEvent) => void | Promise<void>): () => void {
    this.willExecuteListeners.add(listener);
    return () => this.willExecuteListeners.delete(listener);
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

  private async emitWillExecute(evt: CommandWillExecuteEvent): Promise<void> {
    for (const listener of [...this.willExecuteListeners]) {
      await listener(evt);
    }
  }

  private emitDidExecute(evt: CommandExecutionEvent): void {
    for (const listener of [...this.executeListeners]) {
      try {
        listener(evt);
      } catch {
        // ignore
      }
    }
  }

  registerBuiltinCommand(
    commandId: string,
    title: string,
    run: (...args: any[]) => any,
    options?: {
      category?: string | null;
      icon?: string | null;
      description?: string | null;
      keywords?: string[] | null;
      when?: string | null;
    },
  ): void {
    const id = String(commandId);
    const keywords =
      Array.isArray(options?.keywords) && options?.keywords.length > 0
        ? options.keywords
            .map((kw) => (typeof kw === "string" ? kw.trim() : ""))
            .filter((kw) => kw !== "")
        : options?.keywords ?? null;
    this.commands.set(id, {
      commandId: id,
      title,
      category: options?.category ?? null,
      icon: options?.icon ?? null,
      description: options?.description ?? null,
      keywords,
      when: options?.when ?? null,
      source: { kind: "builtin" },
      run: async (...args) => run(...args),
    });
    this.emit();
  }

  /**
   * Replace all extension commands in the registry with the given set.
   */
  setExtensionCommands(
    commands: Array<{
      extensionId: string;
      command: string;
      title: string;
      category?: string | null;
      icon?: string | null;
      description?: string | null;
      keywords?: string[] | null;
      when?: string | null;
    }>,
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
        description: cmd.description ?? null,
        keywords: Array.isArray(cmd.keywords)
          ? cmd.keywords
              .map((kw) => (typeof kw === "string" ? kw.trim() : ""))
              .filter((kw) => kw !== "")
          : null,
        when: cmd.when ?? null,
        source: { kind: "extension", extensionId: String(cmd.extensionId) },
        run: async (...args) => executor(id, ...args),
      });
    }

    this.emit();
  }

  listCommands(): CommandContribution[] {
    return [...this.commands.values()].map(({ run: _run, ...rest }) => rest);
  }

  getCommand(commandId: string): CommandContribution | undefined {
    const entry = this.commands.get(String(commandId));
    if (!entry) return undefined;
    const { run: _run, ...rest } = entry;
    return rest;
  }

  async executeCommand(commandId: string, ...args: any[]): Promise<any> {
    const id = String(commandId);
    const entry = this.commands.get(id);
    if (!entry) throw new Error(`Unknown command: ${commandId}`);
    const { run, ...command } = entry;
    let result: any;
    let error: unknown = null;
    let didThrow = false;
    try {
      await this.emitWillExecute({ commandId: id, args, command });
      result = await run(...args);
      return result;
    } catch (err) {
      didThrow = true;
      error = err;
      throw err;
    } finally {
      this.emitDidExecute({ commandId: id, args, ...(didThrow ? { error } : { result }) });
    }
  }
}
