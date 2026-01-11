import type { ContextKeyLookup, ContextKeyValue } from "./whenClause.js";

export class ContextKeyService {
  private readonly values = new Map<string, ContextKeyValue>();
  private readonly listeners = new Set<() => void>();

  get(key: string): ContextKeyValue {
    return this.values.get(String(key));
  }

  set(key: string, value: ContextKeyValue): void {
    const id = String(key);
    const prev = this.values.get(id);
    if (prev === value) return;
    this.values.set(id, value);
    this.emit();
  }

  /**
   * Convenience: update several keys at once and emit a single change.
   */
  batch(update: Record<string, ContextKeyValue>): void {
    let changed = false;
    for (const [key, value] of Object.entries(update)) {
      const id = String(key);
      const prev = this.values.get(id);
      if (prev === value) continue;
      this.values.set(id, value);
      changed = true;
    }
    if (changed) this.emit();
  }

  onDidChange(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  asLookup(): ContextKeyLookup {
    return (key) => this.get(key);
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
}
