import type { Page } from "@playwright/test";

type CollabSessionStubOptions = {
  sessionId?: string;
  docId?: string;
};

/**
 * Installs a minimal-but-realistic CollabSession stub onto `window.__formulaApp`.
 *
 * This is used by tests that need collab mode semantics (e.g. unsaved prompt suppression)
 * without spinning up a real sync server.
 *
 * The stub implements the pieces of the CollabSession API that desktop `main.ts` touches
 * (notably the `sheets` Y.Array-like surface + `transactLocal`).
 */
export async function installCollabSessionStub(page: Page, options: CollabSessionStubOptions = {}): Promise<void> {
  const { sessionId = "e2e-collab-session", docId = "e2e-doc" } = options;

  await page.evaluate(
    ({ sessionId, docId }) => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing __formulaApp");

      const sheetId = (typeof app.getCurrentSheetId === "function" ? app.getCurrentSheetId() : null) ?? "Sheet1";
      const sheetsState: Array<{ id: string; name: string }> = [{ id: sheetId, name: sheetId }];

      const sheets = {
        toArray: () => sheetsState.map((s) => ({ ...s })),
        observeDeep: () => {},
        unobserveDeep: () => {},
        get length() {
          return sheetsState.length;
        },
        get: (idx: number) => {
          const entry = sheetsState[idx];
          if (!entry) return null;
          return {
            get: (key: string) => (key === "id" ? entry.id : key === "name" ? entry.name : undefined),
            set: (key: string, value: unknown) => {
              if (key === "id") entry.id = String(value ?? "");
              if (key === "name") entry.name = String(value ?? "");
            },
          };
        },
        insert: (index: number, items: any[]) => {
          if (!Array.isArray(items) || items.length === 0) return;
          const normalized: Array<{ id: string; name: string }> = [];
          for (const item of items) {
            const id = String((item as any)?.get?.("id") ?? (item as any)?.id ?? "").trim();
            if (!id) continue;
            const name = String((item as any)?.get?.("name") ?? (item as any)?.name ?? id);
            normalized.push({ id, name });
          }
          sheetsState.splice(index, 0, ...normalized);
        },
        delete: (index: number, count: number) => {
          const n = Number.isFinite(count) ? Math.max(0, Math.trunc(count)) : 0;
          sheetsState.splice(index, n);
        },
      };

      const session = {
        id: sessionId,
        docId,
        sheets,
        transactLocal: (fn: () => void) => fn(),
      };

      app.getCollabSession = () => session;
    },
    { sessionId, docId },
  );
}

