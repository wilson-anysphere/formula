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
      style={{
        position: "absolute",
        inset: 0,
        zIndex: 50,
        background: "var(--dialog-backdrop)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 12,
      }}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) props.onReject();
      }}
    >
      <div
        style={{
          width: "min(560px, 100%)",
          maxHeight: "100%",
          overflow: "auto",
          background: "var(--dialog-bg)",
          border: "1px solid var(--dialog-border)",
          borderRadius: 8,
          boxShadow: "var(--dialog-shadow)",
        }}
      >
        <div style={{ padding: "12px 12px 8px 12px", borderBottom: "1px solid var(--border)" }}>
          <div style={{ fontWeight: 700 }}>Approve AI changes?</div>
          <div style={{ fontSize: 12, opacity: 0.8, marginTop: 4 }}>
            Tool: <span style={{ fontFamily: "monospace" }}>{call.name}</span>
          </div>
        </div>

        <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontSize: 12, opacity: 0.85 }}>
            Summary: {summary.total_changes} changes (creates={summary.creates}, modifies={summary.modifies}, deletes=
            {summary.deletes})
          </div>

          {preview.approval_reasons.length ? (
            <div style={{ fontSize: 12 }}>
              <div style={{ fontWeight: 600, marginBottom: 4 }}>Approval reasons</div>
              <ul style={{ margin: 0, paddingInlineStart: 18 }}>
                {preview.approval_reasons.map((reason) => (
                  <li key={reason}>{reason}</li>
                ))}
              </ul>
            </div>
          ) : null}

          {preview.warnings.length ? (
            <div style={{ fontSize: 12 }}>
              <div style={{ fontWeight: 600, marginBottom: 4 }}>Warnings</div>
              <ul style={{ margin: 0, paddingInlineStart: 18 }}>
                {preview.warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </div>
          ) : null}

          {preview.changes.length ? (
            <div style={{ fontSize: 12 }}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>Cell changes (preview)</div>
              <div style={{ border: "1px solid var(--border)", borderRadius: 6, overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "var(--bg-secondary)" }}>
                      <th style={{ textAlign: "left", padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>
                        Cell
                      </th>
                      <th style={{ textAlign: "left", padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>
                        Type
                      </th>
                      <th style={{ textAlign: "left", padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>
                        Before
                      </th>
                      <th style={{ textAlign: "left", padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>
                        After
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.changes.map((change) => (
                      <tr key={change.cell}>
                        <td style={{ padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>{change.cell}</td>
                        <td style={{ padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>{change.type}</td>
                        <td style={{ padding: "6px 8px", borderBottom: "1px solid var(--border)", opacity: 0.85 }}>
                          <code style={{ whiteSpace: "pre-wrap" }}>{safeStringify(change.before)}</code>
                        </td>
                        <td style={{ padding: "6px 8px", borderBottom: "1px solid var(--border)" }}>
                          <code style={{ whiteSpace: "pre-wrap" }}>{safeStringify(change.after)}</code>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : null}

          <div style={{ fontSize: 12 }}>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>Arguments</div>
            <pre
              style={{
                margin: 0,
                padding: 10,
                background: "var(--bg-secondary)",
                border: "1px solid var(--border)",
                borderRadius: 6,
                overflow: "auto",
              }}
            >
              {safeStringify(call.arguments)}
            </pre>
          </div>
        </div>

        <div
          style={{
            padding: 12,
            borderTop: "1px solid var(--border)",
            display: "flex",
            justifyContent: "flex-end",
            gap: 8,
          }}
        >
          <button type="button" style={{ padding: "8px 12px" }} onClick={props.onReject}>
            Cancel
          </button>
          <button
            type="button"
            style={{ padding: "8px 12px", background: "var(--accent)", color: "var(--text-on-accent)" }}
            onClick={props.onApprove}
          >
            Approve
          </button>
        </div>
      </div>
    </div>
  );
}
