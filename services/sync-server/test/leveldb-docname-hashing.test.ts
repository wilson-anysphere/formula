import assert from "node:assert/strict";
import { Readable } from "node:stream";
import test from "node:test";

import * as Y from "yjs";

import { LeveldbDocNameHashingLayer, sha256Hex } from "../src/leveldb-docname.js";

test("LeveldbDocNameHashingLayer applies sha256(docName) to y-leveldb docName arguments", async () => {
  const docName = "private-doc/name-XYZ";
  const expectedPersisted = sha256Hex(docName);

  const observedDocNames: string[] = [];
  const observedKeys: string[] = [];
  let clock = 0;

  const inner = {
    async flushDocument(persistedName: string) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["flush", persistedName]));
    },
    async getYDoc(persistedName: string) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["getYDoc", persistedName]));
      return { persistedName };
    },
    async storeUpdate(persistedName: string, _update: Uint8Array) {
      observedDocNames.push(persistedName);
      clock += 1;
      observedKeys.push(JSON.stringify(["v1", persistedName, "update", clock]));
      return clock;
    },
    async clearDocument(persistedName: string) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["clearDocument", persistedName]));
    },
    async setMeta(persistedName: string, metaKey: string, _value: unknown) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["setMeta", persistedName, metaKey]));
    },
    async delMeta(persistedName: string, metaKey: string) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["delMeta", persistedName, metaKey]));
    },
    async getMeta(persistedName: string, metaKey: string) {
      observedDocNames.push(persistedName);
      observedKeys.push(JSON.stringify(["getMeta", persistedName, metaKey]));
      return undefined;
    },
    async destroy() {},
  };

  const hashed = new LeveldbDocNameHashingLayer(inner as any, true);

  await hashed.storeUpdate(docName, new Uint8Array([1]));
  await hashed.getYDoc(docName);
  await hashed.flushDocument(docName);
  await hashed.clearDocument(docName);
  await hashed.setMeta(docName, "k", "v");
  await hashed.getMeta(docName, "k");
  await hashed.delMeta(docName, "k");

  assert.ok(observedDocNames.length > 0);
  for (const name of observedDocNames) {
    assert.equal(name, expectedPersisted);
  }
  for (const key of observedKeys) {
    assert.ok(!key.includes(docName));
    assert.ok(key.includes(expectedPersisted));
  }
});

let yLeveldb: typeof import("y-leveldb") | null = null;
try {
  yLeveldb = await import("y-leveldb");
} catch {
  // y-leveldb is an optional dependency (pulled in by y-websocket).
}

if (!yLeveldb) {
  test(
    "LevelDB docName hashing avoids writing raw docName into y-leveldb keys",
    { skip: "y-leveldb not installed" },
    () => {}
  );
} else {
  const { LeveldbPersistence, keyEncoding } = yLeveldb;

  type InMemoryLevelOpts = {
    keyEncoding: typeof keyEncoding;
    valueEncoding: { encode: (v: any) => any; decode: (v: any) => any };
  };

  class InMemoryLevel {
    supports = { clear: true };

    private readonly entries = new Map<string, { key: Buffer; value: any }>();

    constructor(private readonly opts: InMemoryLevelOpts) {}

    private encodeKey(key: any): Buffer {
      return Buffer.isBuffer(key) ? key : this.opts.keyEncoding.encode(key);
    }

    async get(key: any): Promise<any> {
      const encodedKey = this.encodeKey(key);
      const res = this.entries.get(encodedKey.toString("hex"));
      if (!res) {
        const err: NodeJS.ErrnoException & { notFound?: boolean } = new Error(
          "NotFound"
        );
        err.notFound = true;
        throw err;
      }
      return this.opts.valueEncoding.decode(res.value);
    }

    async put(key: any, value: any): Promise<void> {
      const encodedKey = this.encodeKey(key);
      this.entries.set(encodedKey.toString("hex"), {
        key: encodedKey,
        value: this.opts.valueEncoding.encode(value),
      });
    }

    async del(key: any): Promise<void> {
      const encodedKey = this.encodeKey(key);
      this.entries.delete(encodedKey.toString("hex"));
    }

    async batch(
      ops: Array<{ type: string; key: any; value?: any }>
    ): Promise<void> {
      for (const op of ops) {
        if (op.type === "put") {
          await this.put(op.key, op.value);
        } else if (op.type === "del") {
          await this.del(op.key);
        }
      }
    }

    async clear(opts: {
      gte?: any;
      gt?: any;
      lte?: any;
      lt?: any;
    }): Promise<void> {
      const gte = opts.gte ? this.encodeKey(opts.gte) : null;
      const gt = opts.gt ? this.encodeKey(opts.gt) : null;
      const lte = opts.lte ? this.encodeKey(opts.lte) : null;
      const lt = opts.lt ? this.encodeKey(opts.lt) : null;

      for (const [hex, entry] of this.entries.entries()) {
        if (gte && Buffer.compare(entry.key, gte) < 0) continue;
        if (gt && Buffer.compare(entry.key, gt) <= 0) continue;
        if (lte && Buffer.compare(entry.key, lte) > 0) continue;
        if (lt && Buffer.compare(entry.key, lt) >= 0) continue;
        this.entries.delete(hex);
      }
    }

    createReadStream(opts: any = {}): Readable {
      const {
        gte,
        gt,
        lte,
        lt,
        reverse = false,
        limit,
        keys = true,
        values = true,
      } = opts;

      const gteKey = gte ? this.encodeKey(gte) : null;
      const gtKey = gt ? this.encodeKey(gt) : null;
      const lteKey = lte ? this.encodeKey(lte) : null;
      const ltKey = lt ? this.encodeKey(lt) : null;

      const sorted = [...this.entries.values()].sort((a, b) =>
        Buffer.compare(a.key, b.key)
      );
      if (reverse) sorted.reverse();

      const out: any[] = [];
      for (const entry of sorted) {
        if (gteKey && Buffer.compare(entry.key, gteKey) < 0) continue;
        if (gtKey && Buffer.compare(entry.key, gtKey) <= 0) continue;
        if (lteKey && Buffer.compare(entry.key, lteKey) > 0) continue;
        if (ltKey && Buffer.compare(entry.key, ltKey) >= 0) continue;

        const decodedKey = keys ? this.opts.keyEncoding.decode(entry.key) : null;
        const decodedVal = values
          ? this.opts.valueEncoding.decode(entry.value)
          : null;

        if (keys && values) out.push({ key: decodedKey, value: decodedVal });
        else if (keys) out.push(decodedKey);
        else if (values) out.push(decodedVal);

        if (typeof limit === "number" && out.length >= limit) break;
      }

      return Readable.from(out, { objectMode: true });
    }

    async close(): Promise<void> {}
  }

  async function collectKeys(db: any): Promise<any[]> {
    return await new Promise((resolve, reject) => {
      const keys: any[] = [];
      db.createReadStream({ keys: true, values: false })
        .on("data", (data: any) => keys.push(data))
        .on("error", reject)
        .on("end", () => resolve(keys));
    });
  }

  test(
    "LevelDB docName hashing avoids writing raw docName into y-leveldb keys",
    async () => {
      const docName = "private-doc/name-XYZ";
      const expectedPersisted = sha256Hex(docName);

      let db: any | undefined;
      const level = (_location: string, options: any) => {
        db ??= new InMemoryLevel(options);
        return db;
      };

      const ldb1 = new LeveldbPersistence("mem", { level });
      const hashed1 = new LeveldbDocNameHashingLayer(ldb1 as any, true);

      const ydoc = new Y.Doc();
      ydoc.getText("t").insert(0, "hello");

      await hashed1.storeUpdate(docName, Y.encodeStateAsUpdate(ydoc));
      await hashed1.flushDocument(docName);

      assert.ok(db, "expected underlying level instance to be created");
      const keys = await collectKeys(db);
      assert.ok(keys.length > 0, "expected LevelDB to contain keys");

      for (const entry of keys) {
        const key = Array.isArray(entry) ? entry : entry.key;
        assert.ok(Array.isArray(key), "expected decoded LevelDB key array");

        // Ensure the external docName doesn't appear in any key material.
        assert.ok(!JSON.stringify(key).includes(docName));
        const encodedKey = keyEncoding.encode(key);
        assert.ok(!encodedKey.includes(Buffer.from(docName)));
      }

      // Sanity check that we actually stored under the hashed name.
      assert.ok(
        keys.some((entry) => {
          const key = Array.isArray(entry) ? entry : entry.key;
          return Array.isArray(key) && key.includes(expectedPersisted);
        }),
        "expected at least one key to include the persisted (hashed) docName"
      );

      // Restart: create a new persistence wrapper and ensure we can read data back.
      const ldb2 = new LeveldbPersistence("mem", { level });
      const hashed2 = new LeveldbDocNameHashingLayer(ldb2 as any, true);
      const loaded = (await hashed2.getYDoc(docName)) as Y.Doc;

      assert.equal(loaded.getText("t").toString(), "hello");
    }
  );
}
