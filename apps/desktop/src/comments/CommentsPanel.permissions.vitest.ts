// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { CommentsPanel } from "./CommentsPanel";

afterEach(() => {
  document.body.innerHTML = "";
  // React 18 act env flag is set per-test in `renderCommentsPanel`.
  delete (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
  vi.restoreAllMocks();
});

function renderCommentsPanel(opts: { canComment: boolean }) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);

  act(() => {
    root.render(
      React.createElement(CommentsPanel, {
        cellRef: "A1",
        canComment: opts.canComment,
        currentUser: { id: "me", name: "Me" },
        comments: [
          {
            id: "c1",
            cellRef: "A1",
            kind: "threaded",
            author: { id: "u1", name: "Alice" },
            createdAt: 1,
            updatedAt: 1,
            resolved: false,
            content: "Top-level comment",
            mentions: [],
            replies: [
              {
                id: "r1",
                author: { id: "u2", name: "Bob" },
                createdAt: 2,
                updatedAt: 2,
                content: "Reply",
                mentions: [],
              },
            ],
          },
        ],
        onAddComment: vi.fn(),
        onAddReply: vi.fn(),
        onSetResolved: vi.fn(),
      }),
    );
  });

  return { container, root };
}

describe("CommentsPanel permissions", () => {
  it("disables the composer/actions when canComment=false (viewer role)", () => {
    const { container, root } = renderCommentsPanel({ canComment: false });

    // Existing comments are still visible for viewers.
    expect(container.textContent).toContain("Top-level comment");
    expect(container.textContent).toContain("Reply");

    const resolveButton = container.querySelector<HTMLButtonElement>("button.comment-thread__resolve-button");
    expect(resolveButton).toBeInstanceOf(HTMLButtonElement);
    expect(resolveButton!.disabled).toBe(true);

    const replyInput = container.querySelector<HTMLInputElement>("input.comment-thread__reply-input");
    expect(replyInput).toBeInstanceOf(HTMLInputElement);
    expect(replyInput!.disabled).toBe(true);

    const replySubmit = container.querySelector<HTMLButtonElement>("button.comment-thread__submit-reply-button");
    expect(replySubmit).toBeInstanceOf(HTMLButtonElement);
    expect(replySubmit!.disabled).toBe(true);

    const newCommentInput = container.querySelector<HTMLInputElement>("input.comments-panel__new-comment-input");
    expect(newCommentInput).toBeInstanceOf(HTMLInputElement);
    expect(newCommentInput!.disabled).toBe(true);

    const newCommentSubmit = container.querySelector<HTMLButtonElement>("button.comments-panel__submit-button");
    expect(newCommentSubmit).toBeInstanceOf(HTMLButtonElement);
    expect(newCommentSubmit!.disabled).toBe(true);

    const hint = container.querySelector<HTMLElement>(".comments-panel__readonly-hint");
    expect(hint).toBeInstanceOf(HTMLElement);

    act(() => root.unmount());
  });
});

