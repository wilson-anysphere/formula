import { describe, expect, it, beforeEach } from "vitest";

import {
  COLLAB_USER_STORAGE_KEY,
  __resetCollabUserIdentityForTests,
  getCollabUserIdentity,
} from "../userIdentity";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

describe("collab user identity", () => {
  beforeEach(() => {
    __resetCollabUserIdentityForTests();
  });

  it("is stable across repeated calls with the same localStorage", () => {
    const storage = createInMemoryLocalStorage();
    const a = getCollabUserIdentity({ storage, search: "" });
    const b = getCollabUserIdentity({ storage, search: "" });
    expect(b).toEqual(a);
    expect(storage.getItem(COLLAB_USER_STORAGE_KEY)).toBeTruthy();
  });

  it("respects URL overrides without overwriting stored identity", () => {
    const storage = createInMemoryLocalStorage();
    storage.setItem(
      COLLAB_USER_STORAGE_KEY,
      JSON.stringify({ id: "stored-id", name: "Stored", color: "#000000" }),
    );

    const identity = getCollabUserIdentity({
      storage,
      search: "?collabUserId=override-id&collabUserName=Alice&collabUserColor=%23ff0000",
    });

    expect(identity).toEqual({ id: "override-id", name: "Alice", color: "#ff0000" });

    const storedAfter = JSON.parse(storage.getItem(COLLAB_USER_STORAGE_KEY) ?? "{}") as any;
    expect(storedAfter.id).toBe("stored-id");
  });
});

