import { describe, expect, it } from "vitest";
import { loadConfig } from "../config";

function keyBytes(value: number): Buffer {
  return Buffer.alloc(32, value);
}

describe("secret store keyring config", () => {
  it("loads SECRET_STORE_KEYS in comma-separated form and uses the last entry as current", () => {
    const v1 = keyBytes(0x11);
    const v2 = keyBytes(0x22);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS: `v1:${v1.toString("base64")},v2:${v2.toString("base64")}`
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("v2");
    expect(config.secretStoreKeys.keys.v1).toEqual(v1);
    expect(config.secretStoreKeys.keys.v2).toEqual(v2);
  });

  it("prefers SECRET_STORE_KEYS_JSON over SECRET_STORE_KEYS", () => {
    const json1 = keyBytes(0x33);
    const json2 = keyBytes(0x44);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS: `legacy1:${keyBytes(0xaa).toString("base64")},legacy2:${keyBytes(0xbb).toString("base64")}`,
      SECRET_STORE_KEYS_JSON: JSON.stringify({
        currentKeyId: "json2",
        keys: {
          json1: json1.toString("base64"),
          json2: json2.toString("base64")
        }
      })
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("json2");
    expect(config.secretStoreKeys.keys).toEqual({ json1, json2 });
  });

  it("loads SECRET_STORE_KEYS_JSON in { currentKeyId, keys } format", () => {
    const v1 = keyBytes(1);
    const v2 = keyBytes(2);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS_JSON: JSON.stringify({
        currentKeyId: "v2",
        keys: {
          v1: v1.toString("base64"),
          v2: v2.toString("base64")
        }
      })
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("v2");
    expect(config.secretStoreKeys.keys.v1).toEqual(v1);
    expect(config.secretStoreKeys.keys.v2).toEqual(v2);
  });

  it("loads SECRET_STORE_KEYS_JSON in { current, keys } format", () => {
    const v1 = keyBytes(3);
    const v2 = keyBytes(4);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS_JSON: JSON.stringify({
        current: "v2",
        keys: {
          v1: v1.toString("base64"),
          v2: v2.toString("base64")
        }
      })
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("v2");
    expect(config.secretStoreKeys.keys.v1).toEqual(v1);
    expect(config.secretStoreKeys.keys.v2).toEqual(v2);
  });

  it("defaults SECRET_STORE_KEYS_JSON current key id to the last entry for direct maps", () => {
    const v1 = keyBytes(5);
    const v2 = keyBytes(6);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS_JSON: JSON.stringify({
        v1: v1.toString("base64"),
        v2: v2.toString("base64")
      })
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("v2");
    expect(config.secretStoreKeys.keys.v1).toEqual(v1);
    expect(config.secretStoreKeys.keys.v2).toEqual(v2);
  });

  it("defaults SECRET_STORE_KEYS_JSON current key id to the last entry for arrays", () => {
    const v1 = keyBytes(7);
    const v2 = keyBytes(8);

    const config = loadConfig({
      NODE_ENV: "test",
      SECRET_STORE_KEYS_JSON: JSON.stringify([
        { id: "v1", key: v1.toString("base64") },
        { id: "v2", key: v2.toString("base64") }
      ])
    } as any);

    expect(config.secretStoreKeys.currentKeyId).toBe("v2");
    expect(config.secretStoreKeys.keys.v1).toEqual(v1);
    expect(config.secretStoreKeys.keys.v2).toEqual(v2);
  });

  it("rejects duplicate key ids in SECRET_STORE_KEYS", () => {
    const v1 = keyBytes(0x55);
    expect(() =>
      loadConfig({
        NODE_ENV: "test",
        SECRET_STORE_KEYS: `dup:${v1.toString("base64")},dup:${v1.toString("base64")}`
      } as any)
    ).toThrow(/duplicate/i);
  });
});
