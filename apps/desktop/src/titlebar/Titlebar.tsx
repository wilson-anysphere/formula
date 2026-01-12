import React from "react";

import "./titlebar.css";

export type TitlebarAction = {
  /**
   * Visible label for the action button (e.g. "Share").
   */
  label: string;
  /**
   * Accessible label for screen readers.
   */
  ariaLabel: string;
  /**
   * Optional click handler. When omitted, the action is rendered as a no-op
   * placeholder (useful for mockups / non-integrated UI).
   */
  onClick?: () => void;
  /**
   * Optional visual treatment.
   */
  variant?: "default" | "primary";
};

export type TitlebarProps = {
  /**
   * App / product name shown in the titlebar.
   * Defaults to "Formula".
   */
  appName?: string;
  /**
   * Current document name shown next to the app name.
   * Defaults to "Untitled.xlsx".
   */
  documentName?: string;
  /**
   * Optional quick access toolbar state/actions (e.g. Undo/Redo).
   */
  undoRedo?: {
    canUndo: boolean;
    canRedo: boolean;
    undoLabel: string | null;
    redoLabel: string | null;
    onUndo?: () => void;
    onRedo?: () => void;
  };
  /**
   * Optional actions rendered on the right side of the titlebar.
   */
  actions?: TitlebarAction[];
  /**
   * Optional extra class name for the root element.
   */
  className?: string;
};

const defaultActions: TitlebarAction[] = [
  { label: "Comments", ariaLabel: "Open comments" },
  { label: "Share", ariaLabel: "Share document", variant: "primary" },
];

export function Titlebar({
  appName = "Formula",
  documentName = "Untitled.xlsx",
  undoRedo,
  actions = defaultActions,
  className,
}: TitlebarProps) {
  const undoTitle = undoRedo?.undoLabel ? `Undo ${undoRedo.undoLabel}` : "Undo";
  const redoTitle = undoRedo?.redoLabel ? `Redo ${undoRedo.redoLabel}` : "Redo";

  return (
    <div
      className={["formula-titlebar", "formula-titlebar--component", className].filter(Boolean).join(" ")}
      role="banner"
      aria-label="Titlebar"
    >
      <div className="formula-titlebar__window-controls" aria-label="Window controls">
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--close"
          aria-label="Close window"
        />
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--minimize"
          aria-label="Minimize window"
        />
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--maximize"
          aria-label="Maximize window"
        />
      </div>

      {undoRedo ? (
        <div className="formula-titlebar__quick-access" role="toolbar" aria-label="Quick access toolbar">
          <button
            type="button"
            className="formula-titlebar__quick-access-button"
            data-testid="undo"
            aria-label={undoTitle}
            title={undoTitle}
            disabled={!undoRedo.canUndo}
            onClick={undoRedo.onUndo}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path
                d="M9 14 4 9l5-5"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
              <path
                d="M4 9h10a6 6 0 0 1 0 12h-2"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
          <button
            type="button"
            className="formula-titlebar__quick-access-button"
            data-testid="redo"
            aria-label={redoTitle}
            title={redoTitle}
            disabled={!undoRedo.canRedo}
            onClick={undoRedo.onRedo}
          >
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path
                d="M15 14 20 9l-5-5"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
              <path
                d="M20 9H10a6 6 0 0 0 0 12h2"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>
      ) : null}

      {/* Only the middle section is marked as draggable so controls remain clickable. */}
      <div className="formula-titlebar__drag-region" data-tauri-drag-region>
        <div className="formula-titlebar__titles">
          <span className="formula-titlebar__app-name">{appName}</span>
          <span className="formula-titlebar__document-name" title={documentName}>
            {documentName}
          </span>
        </div>
      </div>

      <div className="formula-titlebar__actions" role="toolbar" aria-label="Titlebar actions">
        {actions.map((action) => {
          const variantClass =
            action.variant === "primary" ? "formula-titlebar__action-button--primary" : "";
          return (
            <button
              key={`${action.ariaLabel}:${action.label}`}
              type="button"
              className={["formula-titlebar__action-button", variantClass].filter(Boolean).join(" ")}
              aria-label={action.ariaLabel}
              onClick={action.onClick}
            >
              {action.label}
            </button>
          );
        })}
      </div>
    </div>
  );
}
