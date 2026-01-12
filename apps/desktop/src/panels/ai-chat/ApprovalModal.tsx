import React, { useEffect, useRef } from "react";

import type { LLMToolCall } from "../../../../../packages/ai-tools/src/llm/integration.js";
import type { ToolPlanPreview } from "../../../../../packages/ai-tools/src/preview/preview-engine.js";

export interface ApprovalModalProps {
  request: { call: LLMToolCall; preview: ToolPlanPreview };
  onApprove: () => void;
  onReject: () => void;
}

function safeStringify(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

export function ApprovalModal(props: ApprovalModalProps): React.ReactElement {
  const { call, preview } = props.request;
  const summary = preview.summary;
  const dialogRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    dialogRef.current?.focus();
  }, []);

  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        event.preventDefault();
        props.onReject();
      }
    }
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [props.onReject]);

  return (
    <div
      role="dialog"
      aria-modal="true"
      tabIndex={-1}
      ref={dialogRef}
      className="ai-chat-approval-modal"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) props.onReject();
      }}
    >
      <div
        className="ai-chat-approval-modal__panel"
      >
        <div className="ai-chat-approval-modal__header">
          <div className="ai-chat-approval-modal__title">Approve AI changes?</div>
          <div className="ai-chat-approval-modal__tool">
            Tool: <span className="ai-chat-approval-modal__tool-name">{call.name}</span>
          </div>
        </div>

        <div className="ai-chat-approval-modal__body">
          <div className="ai-chat-approval-modal__summary">
            Summary: {summary.total_changes} changes (creates={summary.creates}, modifies={summary.modifies}, deletes=
            {summary.deletes})
          </div>

          {preview.approval_reasons.length ? (
            <div className="ai-chat-approval-modal__section">
              <div className="ai-chat-approval-modal__section-title">Approval reasons</div>
              <ul className="ai-chat-approval-modal__list">
                {preview.approval_reasons.map((reason) => (
                  <li key={reason}>{reason}</li>
                ))}
              </ul>
            </div>
          ) : null}

          {preview.warnings.length ? (
            <div className="ai-chat-approval-modal__section">
              <div className="ai-chat-approval-modal__section-title">Warnings</div>
              <ul className="ai-chat-approval-modal__list">
                {preview.warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </div>
          ) : null}

          {preview.changes.length ? (
            <div className="ai-chat-approval-modal__section">
              <div className="ai-chat-approval-modal__changes-title">Cell changes (preview)</div>
              <div className="ai-chat-approval-modal__table-wrap">
                <table className="ai-chat-approval-modal__table">
                  <thead>
                    <tr>
                      <th>Cell</th>
                      <th>Type</th>
                      <th>Before</th>
                      <th>After</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.changes.map((change) => (
                      <tr key={change.cell}>
                        <td>{change.cell}</td>
                        <td>{change.type}</td>
                        <td className="ai-chat-approval-modal__before">
                          <code className="ai-chat-approval-modal__code">{safeStringify(change.before)}</code>
                        </td>
                        <td>
                          <code className="ai-chat-approval-modal__code">{safeStringify(change.after)}</code>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : null}

          <div className="ai-chat-approval-modal__section">
            <div className="ai-chat-approval-modal__section-title">Arguments</div>
            <pre className="ai-chat-approval-modal__args-pre">
              {safeStringify(call.arguments)}
            </pre>
          </div>
        </div>

        <div className="ai-chat-approval-modal__footer">
          <button type="button" className="ai-chat-approval-modal__button" onClick={props.onReject}>
            Cancel
          </button>
          <button
            type="button"
            className="ai-chat-approval-modal__button ai-chat-approval-modal__button--approve"
            onClick={props.onApprove}
          >
            Approve
          </button>
        </div>
      </div>
    </div>
  );
}
