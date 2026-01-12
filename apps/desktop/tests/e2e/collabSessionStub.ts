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
      // Keep a lightweight in-memory representation of the collab workbook sheet list.
      // Use a generic record to support metadata keys beyond {id,name} (visibility, tabColor, etc).
      const sheetsState: Array<Record<string, any>> = [{ id: sheetId, name: sheetId, visibility: "visible" }];

      // `main.ts` observes deep Yjs events on `session.sheets` to keep the desktop sheet UI in sync.
      // Implement a minimal observer registry so tests can emulate collab mode without a real server.
      const observers = new Set<(events: any, transaction: any) => void>();
      const notifyObservers = () => {
        for (const observer of [...observers]) {
          try {
            observer([], { local: true });
          } catch {
            // ignore observer errors
          }
        }
      };

      const getFromYjsPrelim = (item: any, key: string): unknown => {
        // When `main.ts` creates a sheet in collab mode it constructs a new `Y.Map()` and
        // sets fields before inserting into `session.sheets`. In real Yjs, insertion into
        // a Y.Doc integrates the map so `map.get("id")` works.
        //
        // In this e2e stub, there is no real Y.Doc, so the map stays in its "preliminary"
        // state and `get()` warns + returns undefined. The values are stored in the internal
        // `_prelimContent` Map; read from it when present.
        const prelim = item?._prelimContent;
        if (!prelim || typeof prelim.get !== "function") return undefined;
        try {
          return prelim.get(key);
        } catch {
          return undefined;
        }
      };

      const readSheetField = (item: any, key: string): unknown => {
        if (!item || typeof item !== "object") return undefined;
        const fromPrelim = getFromYjsPrelim(item, key);
        if (fromPrelim !== undefined) return fromPrelim;
        if (typeof item.get === "function") {
          try {
            const value = item.get(key);
            if (value !== undefined) return value;
          } catch {
            // ignore
          }
        }
        return (item as any)[key];
      };

      const coerceSheetEntry = (item: any): Record<string, any> | null => {
        if (!item || typeof item !== "object") return null;
        const out: Record<string, any> = {};

        const prelim = item?._prelimContent;
        if (prelim && typeof prelim.forEach === "function") {
          try {
            prelim.forEach((value: unknown, key: string) => {
              out[String(key)] = value;
            });
          } catch {
            // ignore
          }
        } else if (typeof item.forEach === "function") {
          // For our in-memory wrapper objects we can support `forEach` to behave like Y.Map.
          try {
            item.forEach((value: unknown, key: string) => {
              out[String(key)] = value;
            });
          } catch {
            // ignore
          }
        } else {
          for (const [k, v] of Object.entries(item)) {
            out[k] = v;
          }
        }

        const id = String(out.id ?? readSheetField(item, "id") ?? "").trim();
        if (!id) return null;
        const name = String(out.name ?? readSheetField(item, "name") ?? id).trim() || id;

        out.id = id;
        out.name = name;
        if (out.visibility == null) out.visibility = "visible";

        return out;
      };

      const sheets = {
        toArray: () => sheetsState.map((s) => ({ ...s })),
        observeDeep: (handler: any) => {
          if (typeof handler !== "function") return;
          observers.add(handler);
        },
        unobserveDeep: (handler: any) => {
          observers.delete(handler);
        },
        get length() {
          return sheetsState.length;
        },
        get: (idx: number) => {
          const entry = sheetsState[idx];
          if (!entry) return null;
          return {
            get: (key: string) => (entry as any)[key],
            set: (key: string, value: unknown) => {
              (entry as any)[key] = value;
              notifyObservers();
            },
            delete: (key: string) => {
              delete (entry as any)[key];
              notifyObservers();
            },
            has: (key: string) => Object.prototype.hasOwnProperty.call(entry, key),
            forEach: (fn: any) => {
              if (typeof fn !== "function") return;
              for (const [k, v] of Object.entries(entry)) {
                fn(v, k);
              }
            },
            observeDeep: () => {},
            unobserveDeep: () => {},
          };
        },
        insert: (index: number, items: any[]) => {
          if (!Array.isArray(items) || items.length === 0) return;
          const normalized: Array<Record<string, any>> = [];
          for (const item of items) {
            const entry = coerceSheetEntry(item);
            if (!entry) continue;
            normalized.push(entry);
          }
          if (normalized.length === 0) return;
          const idx = Number.isFinite(index) ? Math.max(0, Math.trunc(index)) : 0;
          sheetsState.splice(idx, 0, ...normalized);
          notifyObservers();
        },
        push: (items: any[]) => {
          if (!Array.isArray(items) || items.length === 0) return;
          (sheets as any).insert(sheetsState.length, items);
        },
        delete: (index: number, count: number) => {
          const n = Number.isFinite(count) ? Math.max(0, Math.trunc(count)) : 0;
          sheetsState.splice(index, n);
          if (n > 0) notifyObservers();
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
