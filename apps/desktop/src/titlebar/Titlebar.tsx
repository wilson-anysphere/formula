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
  actions = defaultActions,
  className,
}: TitlebarProps) {
  return (
    <div className={["formula-titlebar", className].filter(Boolean).join(" ")}>
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

      {/* Only the middle section is marked as draggable so controls remain clickable. */}
      <div className="formula-titlebar__drag-region" data-tauri-drag-region>
        <div className="formula-titlebar__titles">
          <span className="formula-titlebar__app-name">{appName}</span>
          <span className="formula-titlebar__document-name" title={documentName}>
            {documentName}
          </span>
        </div>
      </div>

      <div className="formula-titlebar__actions" aria-label="Titlebar actions">
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

