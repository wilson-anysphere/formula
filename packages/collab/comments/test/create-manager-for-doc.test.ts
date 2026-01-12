import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { createCommentManagerForDoc } from "../src/manager";

describe("createCommentManagerForDoc", () => {
  it("does not instantiate the comments root when it doesn't exist yet", () => {
    const doc = new Y.Doc();
    expect(doc.share.get("comments")).toBe(undefined);

    createCommentManagerForDoc({ doc, transact: (fn) => doc.transact(fn) });

    // Creating the manager should be safe to do pre-hydration and must not
    // instantiate `doc.getMap("comments")` eagerly (which could clobber legacy
    // Array-backed docs when remote updates later arrive).
    expect(doc.share.get("comments")).toBe(undefined);
  });
});

