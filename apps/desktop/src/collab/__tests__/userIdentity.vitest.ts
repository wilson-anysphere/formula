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

  it("accepts legacy userId/userName/userColor query params", () => {
    const storage = createInMemoryLocalStorage();
    const identity = getCollabUserIdentity({
      storage,
      search: "?userId=legacy-id&userName=Legacy&userColor=%2300ff00",
    });

    expect(identity).toEqual({ id: "legacy-id", name: "Legacy", color: "#00ff00" });
  });

  it("falls back to generating a new identity when stored data is corrupt", () => {
    const storage = createInMemoryLocalStorage();
    storage.setItem(COLLAB_USER_STORAGE_KEY, "not-json");

    const identity = getCollabUserIdentity({
      storage,
      search: "",
      crypto: { randomUUID: () => "uuid-0000" } as any,
    });

    expect(identity.id).toBe("uuid-0000");
    expect(identity.name).toBe("User uuid");
    expect(identity.color).toMatch(/^#[0-9a-f]{6}$/);
    expect(storage.getItem(COLLAB_USER_STORAGE_KEY)).toContain("uuid-0000");
  });

  it("chooses a deterministic default name + color from a provided collabUserId", () => {
    const a = getCollabUserIdentity({ storage: createInMemoryLocalStorage(), search: "?collabUserId=abc123" });
    const b = getCollabUserIdentity({ storage: createInMemoryLocalStorage(), search: "?collabUserId=abc123" });
    expect(a).toEqual(b);
    expect(a.id).toBe("abc123");
    expect(a.name).toBe("User abc1");
    expect(a.color).toMatch(/^#[0-9a-f]{6}$/);
  });

  it("falls back to an in-memory identity when storage throws", () => {
    const throwingStorage = {
      getItem: () => {
        throw new Error("storage disabled");
      },
      setItem: () => {
        throw new Error("storage disabled");
      },
      removeItem: () => {},
      clear: () => {},
      key: () => null,
      length: 0,
    } as unknown as Storage;

    const identityA = getCollabUserIdentity({
      storage: throwingStorage,
      search: "",
      crypto: { randomUUID: () => "uuid-throw" } as any,
    });
    const identityB = getCollabUserIdentity({
      storage: throwingStorage,
      search: "",
      // Prove we don't re-generate a new id on the second call even if crypto differs.
      crypto: { randomUUID: () => "uuid-other" } as any,
    });

    expect(identityA).toEqual({ id: "uuid-throw", name: "User uuid", color: identityA.color });
    expect(identityA.color).toMatch(/^#[0-9a-f]{6}$/);
    expect(identityB).toEqual(identityA);
  });
});
