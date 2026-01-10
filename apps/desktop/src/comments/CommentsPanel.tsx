import React, { useMemo, useState } from "react";

import type { Comment, CommentAuthor } from "@formula/collab-comments";
import { t, tWithVars } from "../i18n/index.js";

export interface CommentsPanelProps {
  cellRef: string | null;
  comments: Comment[];
  currentUser: CommentAuthor;
  onAddComment: (input: { cellRef: string; content: string }) => void;
  onAddReply: (input: { commentId: string; content: string }) => void;
  onSetResolved: (input: { commentId: string; resolved: boolean }) => void;
}

export function CommentsPanel(props: CommentsPanelProps): React.ReactElement {
  const [newComment, setNewComment] = useState("");
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});

  const threads = useMemo(() => {
    return props.comments
      .slice()
      .sort((a, b) => a.createdAt - b.createdAt);
  }, [props.comments]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ padding: 12, borderBottom: "1px solid var(--border)" }}>
        <div style={{ fontWeight: 600 }}>{t("comments.title")}</div>
        <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>
          {props.cellRef ? tWithVars("comments.cellLabel", { cellRef: props.cellRef }) : t("comments.selectCell")}
        </div>
      </div>

      <div style={{ flex: 1, overflow: "auto", padding: 12, gap: 12, display: "flex", flexDirection: "column" }}>
        {threads.length === 0 ? (
          <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("comments.none")}</div>
        ) : (
          threads.map((comment) => (
            <div key={comment.id} style={{ border: "1px solid var(--border)", borderRadius: 8, padding: 10 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                <div style={{ fontSize: 12, fontWeight: 600 }}>{comment.author.name}</div>
                <button
                  type="button"
                  onClick={() => props.onSetResolved({ commentId: comment.id, resolved: !comment.resolved })}
                >
                  {comment.resolved ? t("comments.unresolve") : t("comments.resolve")}
                </button>
              </div>

              <div style={{ fontSize: 13, marginTop: 6, whiteSpace: "pre-wrap" }}>{comment.content}</div>

              <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 8 }}>
                {comment.replies
                  .slice()
                  .sort((a, b) => a.createdAt - b.createdAt)
                  .map((reply) => (
                    <div
                      key={reply.id}
                      style={{ paddingInlineStart: 12, borderInlineStart: "2px solid var(--border)" }}
                    >
                      <div style={{ fontSize: 12, fontWeight: 600 }}>{reply.author.name}</div>
                      <div style={{ fontSize: 13, marginTop: 4, whiteSpace: "pre-wrap" }}>{reply.content}</div>
                    </div>
                  ))}

                <div style={{ display: "flex", gap: 8 }}>
                  <input
                    value={replyDrafts[comment.id] ?? ""}
                    onChange={(e) =>
                      setReplyDrafts((drafts) => ({
                        ...drafts,
                        [comment.id]: e.target.value,
                      }))
                    }
                    placeholder={t("comments.reply.placeholder")}
                    style={{ flex: 1 }}
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const draft = (replyDrafts[comment.id] ?? "").trim();
                      if (!draft) return;
                      props.onAddReply({ commentId: comment.id, content: draft });
                      setReplyDrafts((drafts) => ({ ...drafts, [comment.id]: "" }));
                    }}
                  >
                    {t("comments.reply.send")}
                  </button>
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      <div style={{ padding: 12, borderTop: "1px solid var(--border)" }}>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            value={newComment}
            onChange={(e) => setNewComment(e.target.value)}
            placeholder={t("comments.new.placeholder")}
            style={{ flex: 1 }}
            disabled={!props.cellRef}
          />
          <button
            type="button"
            disabled={!props.cellRef || newComment.trim().length === 0}
            onClick={() => {
              if (!props.cellRef) return;
              const content = newComment.trim();
              if (!content) return;
              props.onAddComment({ cellRef: props.cellRef, content });
              setNewComment("");
            }}
          >
            {t("comments.new.submit")}
          </button>
        </div>
      </div>
    </div>
  );
}
