import React from "react";

import { RedoIcon, UndoIcon } from "../ui/icons/index.js";

import "./titlebar.css";

export type TitlebarAction = {
  /**
   * Optional stable identifier used as the React `key` when rendering the action list.
   *
   * Useful when actions have the same label/aria-label but represent distinct commands.
   */
  id?: string;
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
  /**
   * Optional disabled state (useful for placeholders).
   */
  disabled?: boolean;
};

export type TitlebarWindowControls = {
  onClose?: () => void;
  onMinimize?: () => void;
  onToggleMaximize?: () => void;
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
   * Optional callbacks for native desktop window controls (Tauri desktop builds).
   *
   * When a callback is omitted, the corresponding window control is rendered as
   * disabled to make it clear the action isn't available.
   */
  windowControls?: TitlebarWindowControls;
  /**
   * Optional extra class name for the root element.
   */
  className?: string;
};

const defaultActions: TitlebarAction[] = [
  { label: "Comments", ariaLabel: "Open comments" },
  { label: "Share", ariaLabel: "Share document", variant: "primary" },
];

function stripLeadingDocSeparator(documentName: string): string {
  // The titlebar visually separates app name and document name with an em dash.
  // Some callers may already include a leading dash in `documentName` (e.g. from
  // mockups). Strip it so we don't render a doubled separator.
  let result = documentName.trimStart();
  // Remove any leading separator characters and following whitespace.
  // (em dash, en dash, hyphen-minus)
  result = result.replace(/^[—–-]+\s*/, "");
  return result;
}

export function Titlebar({
  appName = "Formula",
  documentName = "Untitled.xlsx",
  undoRedo,
  actions = defaultActions,
  windowControls,
  className,
}: TitlebarProps) {
  const undoTitle = undoRedo?.undoLabel ? `Undo ${undoRedo.undoLabel}` : "Undo";
  const redoTitle = undoRedo?.redoLabel ? `Redo ${undoRedo.redoLabel}` : "Redo";
  const normalizedDocumentName = stripLeadingDocSeparator(documentName);
  const canUndo = Boolean(undoRedo?.canUndo) && typeof undoRedo?.onUndo === "function";
  const canRedo = Boolean(undoRedo?.canRedo) && typeof undoRedo?.onRedo === "function";

  return (
    <div
      className={["formula-titlebar", "formula-titlebar--component", className].filter(Boolean).join(" ")}
      role="banner"
      aria-label="Titlebar"
      data-testid="titlebar-component"
    >
      <div
        className="formula-titlebar__window-controls"
        role="group"
        aria-label="Window controls"
        data-testid="titlebar-window-controls"
      >
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--close"
          aria-label="Close window"
          title="Close window"
          data-testid="titlebar-window-close"
          disabled={!windowControls?.onClose}
          onClick={windowControls?.onClose}
        />
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--minimize"
          aria-label="Minimize window"
          title="Minimize window"
          data-testid="titlebar-window-minimize"
          disabled={!windowControls?.onMinimize}
          onClick={windowControls?.onMinimize}
        />
        <button
          type="button"
          className="formula-titlebar__window-button formula-titlebar__window-button--maximize"
          aria-label="Maximize window"
          title="Maximize window"
          data-testid="titlebar-window-maximize"
          disabled={!windowControls?.onToggleMaximize}
          onClick={windowControls?.onToggleMaximize}
        />
      </div>

      {undoRedo ? (
        <div
          className="formula-titlebar__quick-access"
          role="toolbar"
          aria-label="Quick access toolbar"
          data-testid="titlebar-quick-access"
        >
          <button
            type="button"
            className="formula-titlebar__quick-access-button"
            data-testid="undo"
            aria-label={undoTitle}
            title={undoTitle}
            disabled={!canUndo}
            onClick={undoRedo.onUndo}
          >
            <UndoIcon />
          </button>
          <button
            type="button"
            className="formula-titlebar__quick-access-button"
            data-testid="redo"
            aria-label={redoTitle}
            title={redoTitle}
            disabled={!canRedo}
            onClick={undoRedo.onRedo}
          >
            <RedoIcon />
          </button>
        </div>
      ) : null}

      {/* Only the middle section is marked as draggable so controls remain clickable. */}
      <div
        className="formula-titlebar__drag-region"
        data-tauri-drag-region
        data-testid="titlebar-drag-region"
        onDoubleClick={windowControls?.onToggleMaximize}
      >
        <div className="formula-titlebar__titles" data-testid="titlebar-titles">
          <span className="formula-titlebar__app-name" data-testid="titlebar-app-name">
            {appName}
          </span>
          {normalizedDocumentName.trim().length > 0 ? (
            <span className="formula-titlebar__document-name" title={normalizedDocumentName} data-testid="titlebar-document-name">
              {normalizedDocumentName}
            </span>
          ) : null}
        </div>
      </div>

      {actions.length > 0 ? (
        <div className="formula-titlebar__actions" role="toolbar" aria-label="Titlebar actions" data-testid="titlebar-actions">
          {actions.map((action) => {
            const variantClass =
              action.variant === "primary" ? "formula-titlebar__action-button--primary" : "";
            return (
              <button
                key={action.id ?? `${action.ariaLabel}:${action.label}`}
                type="button"
                className={["formula-titlebar__action-button", variantClass].filter(Boolean).join(" ")}
                aria-label={action.ariaLabel}
                title={action.ariaLabel}
                onClick={action.onClick}
                disabled={action.disabled}
              >
                {action.label}
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}
