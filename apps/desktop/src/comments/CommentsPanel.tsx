import React, { useMemo, useState } from "react";

import type { Comment, CommentAuthor } from "@formula/collab-comments";
import { t, tWithVars } from "../i18n/index.js";

export interface CommentsPanelProps {
  cellRef: string | null;
  comments: Comment[];
  currentUser: CommentAuthor;
  /**
   * Whether the current user is allowed to create/update comments.
   * Viewers can read existing threads but cannot comment.
   */
  canComment?: boolean;
  onAddComment: (input: { cellRef: string; content: string }) => void;
  onAddReply: (input: { commentId: string; content: string }) => void;
  onSetResolved: (input: { commentId: string; resolved: boolean }) => void;
}

export function CommentsPanel(props: CommentsPanelProps): React.ReactElement {
  const [newComment, setNewComment] = useState("");
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});
  const canComment = props.canComment ?? true;

  const threads = useMemo(() => {
    return props.comments
      .slice()
      .sort((a, b) => a.createdAt - b.createdAt);
  }, [props.comments]);

  return (
    <div className="comments-panel-view">
      <div className="comments-panel-view__header">
        <div className="comments-panel-view__title">{t("comments.title")}</div>
        <div className="comments-panel-view__subtitle">
          {props.cellRef ? tWithVars("comments.cellLabel", { cellRef: props.cellRef }) : t("comments.selectCell")}
        </div>
      </div>

      <div className="comments-panel-view__body">
        {threads.length === 0 ? (
          <div className="comments-panel__empty">{t("comments.none")}</div>
        ) : (
          threads.map((comment) => (
            <div
              key={comment.id}
              className="comment-thread"
              data-resolved={comment.resolved ? "true" : "false"}
            >
              <div className="comment-thread__header">
                <div className="comment-thread__author">{comment.author.name || t("presence.anonymous")}</div>
                <button
                  type="button"
                  className="comment-thread__resolve-button"
                  disabled={!canComment}
                  onClick={() => {
                    if (!canComment) return;
                    props.onSetResolved({ commentId: comment.id, resolved: !comment.resolved });
                  }}
                >
                  {comment.resolved ? t("comments.unresolve") : t("comments.resolve")}
                </button>
              </div>

              <div className="comment-thread__body">{comment.content}</div>

              {comment.replies
                .slice()
                .sort((a, b) => a.createdAt - b.createdAt)
                .map((reply) => (
                  <div key={reply.id} className="comment-thread__reply">
                    <div className="comment-thread__reply-author">{reply.author.name || t("presence.anonymous")}</div>
                    <div className="comment-thread__reply-body">{reply.content}</div>
                  </div>
                ))}

              <div className="comment-thread__reply-row">
                <input
                  value={replyDrafts[comment.id] ?? ""}
                  className="comment-thread__reply-input"
                  disabled={!canComment}
                  onChange={(e) =>
                    setReplyDrafts((drafts) => ({
                      ...drafts,
                      [comment.id]: e.target.value,
                    }))
                  }
                  placeholder={t("comments.reply.placeholder")}
                />
                <button
                  type="button"
                  className="comment-thread__submit-reply-button"
                  disabled={!canComment}
                  onClick={() => {
                    if (!canComment) return;
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
          ))
        )}
      </div>

      <div className="comments-panel-view__footer">
        {!canComment ? <div className="comments-panel__readonly-hint">{t("comments.readOnlyHint")}</div> : null}
        <div className="comments-panel-view__row">
          <input
            value={newComment}
            onChange={(e) => setNewComment(e.target.value)}
            placeholder={t("comments.new.placeholder")}
            className="comments-panel__new-comment-input"
            disabled={!props.cellRef || !canComment}
          />
          <button
            type="button"
            className="comments-panel__submit-button"
            disabled={!props.cellRef || !canComment || newComment.trim().length === 0}
            onClick={() => {
              if (!canComment) return;
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
