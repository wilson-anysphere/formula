import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN, createCollabUndoService } from "../index.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

function replacePlaceholderRootType({ doc, name, existing, create }) {
  const t = create();

  // Mirror Yjs' `Doc.get()` behavior when turning an `AbstractType` placeholder
  // root into a concrete type, but allow using a constructor from another Yjs
  // module instance (ESM vs CJS).
  t._map = existing?._map;
  if (t._map instanceof Map) {
    t._map.forEach((n) => {
      for (; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  t._start = existing?._start;
  for (let n = t._start; n !== null; n = n.right) {
    n.parent = t;
  }
  t._length = existing?._length;

  doc.share.set(name, t);
  t._integrate?.(doc, null);

  return t;
}

test("collab undo: does not warn [yjs#509] when scope contains a foreign root type (CJS constructor)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const root = remote.getMap("comments");
    const comment = new Ycjs.Map();
    comment.set("x", 1);
    root.set("c1", comment);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  const placeholder = doc.share.get("comments");
  assert.ok(placeholder, "expected comments root placeholder to exist");
  assert.equal(placeholder.constructor?.name, "AbstractType", "expected generic root placeholder");

  // Find the nested map value created by the CJS build.
  const item = placeholder._map?.get("c1");
  assert.ok(item, "expected placeholder to contain a map entry for 'c1'");
  const content = item.content?.getContent?.() ?? [];
  const nested = content[content.length - 1];
  assert.ok(nested, "expected nested comment map to exist");
  assert.equal(nested instanceof Y.Map, false, "expected nested comment map to be a foreign Y.Map instance");

  // Replace the placeholder root with a Map created by the same foreign module
  // instance as its nested values (mirrors `getCommentsRoot`).
  const MapCtor = nested.constructor;
  assert.equal(typeof MapCtor, "function", "expected nested comment map to have a constructor");
  const foreignRoot = replacePlaceholderRootType({
    doc,
    name: "comments",
    existing: placeholder,
    create: () => new MapCtor(),
  });
  // Note: the undo service patches foreign prototype chains so foreign types can
  // pass `instanceof Y.AbstractType` checks. Don't assert on that here; instead
  // assert the root is still a foreign Map constructor.
  assert.equal(foreignRoot instanceof Y.Map, false, "expected foreign root map type (not instanceof ESM Y.Map)");

  const warns = [];
  const prevWarn = console.warn;
  console.warn = (...args) => {
    warns.push(args.join(" "));
  };
  try {
    const undo = createCollabUndoService({ doc, scope: foreignRoot });

    // Ensure undo works for edits in the foreign subtree (and that we didn't
    // break the foreign type by patching its constructor prototype chain).
    undo.transact(() => {
      nested.set("x", 2);
    });
    undo.stopCapturing();
    assert.equal(nested.get("x"), 2);

    assert.equal(undo.canUndo(), true);
    undo.undo();
    assert.equal(nested.get("x"), 1);

    undo.undoManager.destroy();
  } finally {
    console.warn = prevWarn;
  }

  assert.equal(
    warns.some((w) => w.includes("[yjs#509] Not same Y.Doc")),
    false,
    `expected no [yjs#509] warning, got:\n${warns.join("\n")}`
  );

  doc.destroy();
  remote.destroy();
});

test("collab undo: does not warn [yjs#509] when adding a foreign type to scope later (UndoManager.addToScope)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const root = remote.getMap("comments");
    const comment = new Ycjs.Map();
    comment.set("x", 1);
    root.set("c1", comment);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const undo = createCollabUndoService({ doc, scope: cells });

  // Hydrate a foreign comments root.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);
  const placeholder = doc.share.get("comments");
  assert.ok(placeholder, "expected comments root placeholder to exist");

  const item = placeholder._map?.get("c1");
  assert.ok(item);
  const content = item.content?.getContent?.() ?? [];
  const nested = content[content.length - 1];
  assert.ok(nested);
  const MapCtor = nested.constructor;
  const foreignRoot = replacePlaceholderRootType({
    doc,
    name: "comments",
    existing: placeholder,
    create: () => new MapCtor(),
  });

  const warns = [];
  const prevWarn = console.warn;
  console.warn = (...args) => {
    warns.push(args.join(" "));
  };
  try {
    undo.undoManager.addToScope(foreignRoot);
  } finally {
    console.warn = prevWarn;
  }

  assert.equal(
    warns.some((w) => w.includes("[yjs#509] Not same Y.Doc")),
    false,
    `expected no [yjs#509] warning, got:\n${warns.join("\n")}`
  );

  undo.undoManager.destroy();
  doc.destroy();
  remote.destroy();
});
