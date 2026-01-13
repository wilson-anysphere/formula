import React, { useMemo, useState } from "react";

import type { ArrowTableAdapter, DataTable, Query, QueryOperation } from "@formula/power-query";
import { t } from "../../../i18n/index.js";

export function AddStepMenu(props: {
  onAddStep: (op: QueryOperation) => void;
  onAiSuggest?: (intent: string, ctx: { query: Query; preview: DataTable | ArrowTableAdapter | null }) => Promise<QueryOperation[]>;
  aiContext: { query: Query; preview: DataTable | ArrowTableAdapter | null };
}) {
  const [intent, setIntent] = useState("");
  const [suggestions, setSuggestions] = useState<QueryOperation[] | null>(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);

  const [menuOpen, setMenuOpen] = useState(false);

  const schema = props.aiContext.preview?.columns ?? [];
  const columnNames = useMemo(() => schema.map((col) => col.name).filter((name) => name.trim().length > 0), [schema]);
  const schemaReady = columnNames.length > 0;
  const firstColumnName = columnNames[0] ?? "";
  const secondColumnName = columnNames[1] ?? firstColumnName;

  type MenuItem = {
    id: string;
    label: string;
    disabled?: boolean;
    disabledReason?: string;
    create: () => QueryOperation;
  };

  const schemaRequiredReason = t("queryEditor.addStep.schemaRequired");

  const menuGroups = useMemo(() => {
    const schemaRequiredDisabled = !schemaReady;
    const schemaItem = (item: Omit<MenuItem, "disabled" | "disabledReason">): MenuItem => ({
      ...item,
      disabled: schemaRequiredDisabled,
      disabledReason: schemaRequiredDisabled ? schemaRequiredReason : undefined,
    });

    return [
      {
        id: "rows",
        label: t("queryEditor.addStep.category.rows"),
        items: [
          schemaItem({
            id: "filterRows",
            label: t("queryEditor.addStep.op.filterRows"),
            create: () => ({
              type: "filterRows",
              predicate: { type: "comparison", column: firstColumnName, operator: "isNotNull" },
            }),
          }),
          schemaItem({
            id: "sortRows",
            label: t("queryEditor.addStep.op.sort"),
            create: () => ({
              type: "sortRows",
              sortBy: [{ column: firstColumnName, direction: "ascending" }],
            }),
          }),
        ] satisfies MenuItem[],
      },
      {
        id: "columns",
        label: t("queryEditor.addStep.category.columns"),
        items: [
          schemaItem({
            id: "removeColumns",
            label: t("queryEditor.addStep.op.removeColumns"),
            create: () => ({ type: "removeColumns", columns: [firstColumnName] }),
          }),
          schemaItem({
            id: "keepColumns",
            label: t("queryEditor.addStep.op.keepColumns"),
            create: () => ({ type: "selectColumns", columns: [firstColumnName] }),
          }),
          schemaItem({
            id: "renameColumn",
            label: t("queryEditor.addStep.op.renameColumns"),
            create: () => ({ type: "renameColumn", oldName: firstColumnName, newName: `${firstColumnName || "Column"} (Renamed)` }),
          }),
          schemaItem({
            id: "splitColumn",
            label: t("queryEditor.addStep.op.splitColumn"),
            create: () => ({ type: "splitColumn", column: firstColumnName, delimiter: "," }),
          }),
        ] satisfies MenuItem[],
      },
      {
        id: "transform",
        label: t("queryEditor.addStep.category.transform"),
        items: [
          schemaItem({
            id: "changeType",
            label: t("queryEditor.addStep.op.changeType"),
            create: () => ({ type: "changeType", column: firstColumnName, newType: "string" }),
          }),
          schemaItem({
            id: "groupBy",
            label: t("queryEditor.addStep.op.groupBy"),
            create: () => ({
              type: "groupBy",
              groupColumns: [firstColumnName],
              aggregations: [{ column: secondColumnName, op: "count", as: t("queryEditor.addStep.op.groupBy.count") }],
            }),
          }),
        ] satisfies MenuItem[],
      },
    ];
  }, [firstColumnName, schemaReady, schemaRequiredReason, secondColumnName]);

  function humanizeOperationType(type: string): string {
    return type
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/^./, (ch) => ch.toUpperCase());
  }

  function describeOperation(op: QueryOperation): string {
    switch (op.type) {
      case "filterRows": {
        const column =
          op.predicate?.type === "comparison" && typeof op.predicate.column === "string" ? op.predicate.column : "";
        return column ? `${t("queryEditor.addStep.op.filterRows")} (${column})` : t("queryEditor.addStep.op.filterRows");
      }
      case "sortRows": {
        const col = op.sortBy?.[0]?.column ?? "";
        return col ? `${t("queryEditor.addStep.op.sort")} (${col})` : t("queryEditor.addStep.op.sort");
      }
      case "removeColumns":
        return t("queryEditor.addStep.op.removeColumns");
      case "selectColumns":
        return t("queryEditor.addStep.op.keepColumns");
      case "renameColumn":
        return t("queryEditor.addStep.op.renameColumns");
      case "changeType":
        return t("queryEditor.addStep.op.changeType");
      case "splitColumn":
        return t("queryEditor.addStep.op.splitColumn");
      case "groupBy":
        return t("queryEditor.addStep.op.groupBy");
      default:
        return humanizeOperationType(op.type);
    }
  }

  return (
    <div className="query-editor-add-step">
      <div className="query-editor-add-step__menu">
        <button
          type="button"
          className="query-editor-add-step__menu-trigger"
          onClick={() => setMenuOpen((open) => !open)}
          aria-haspopup="menu"
          aria-expanded={menuOpen}
        >
          {t("queryEditor.addStep.addStep")}
        </button>
        {menuOpen ? (
          <div className="query-editor-add-step__menu-popover" role="menu">
            {menuGroups.map((group) => (
              <div key={group.id} className="query-editor-add-step__menu-group">
                <div className="query-editor-add-step__menu-group-title">{group.label}</div>
                <div className="query-editor-add-step__menu-group-items">
                  {group.items.map((item) => (
                    <button
                      key={item.id}
                      type="button"
                      role="menuitem"
                      disabled={item.disabled}
                      title={item.disabled ? item.disabledReason : undefined}
                      className="query-editor-add-step__menu-item"
                      onClick={() => {
                        if (item.disabled) return;
                        props.onAddStep(item.create());
                        setMenuOpen(false);
                      }}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>
              </div>
            ))}
            {!schemaReady ? <div className="query-editor-add-step__menu-hint">{schemaRequiredReason}</div> : null}
          </div>
        ) : null}
      </div>

      {props.onAiSuggest ? (
        <div>
          <input
            value={intent}
            onChange={(e) => setIntent(e.target.value)}
            placeholder={t("queryEditor.addStep.aiPlaceholder")}
            className="query-editor-add-step__ai-input"
          />
          <button
            type="button"
            onClick={async () => {
              const trimmed = intent.trim();
              if (!trimmed) return;
              setAiLoading(true);
              setAiError(null);
              setSuggestions(null);
              try {
                const ops = await props.onAiSuggest?.(trimmed, props.aiContext);
                setSuggestions(ops ?? []);
              } catch (err) {
                setAiError(err instanceof Error ? err.message : String(err));
                setSuggestions([]);
              } finally {
                setAiLoading(false);
              }
            }}
            disabled={!intent.trim() || aiLoading}
            className="query-editor-add-step__ai-button"
          >
            {aiLoading ? t("queryEditor.addStep.suggestingNext") : t("queryEditor.addStep.suggestNext")}
          </button>
          {aiError ? <div className="query-editor-add-step__ai-error">{aiError}</div> : null}
          {suggestions ? (
            <div className="query-editor-add-step__suggestions">
              {suggestions.length === 0 ? (
                <div className="query-editor-add-step__no-suggestions">{t("queryEditor.addStep.noSuggestions")}</div>
              ) : (
                suggestions.map((op, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => props.onAddStep(op)}
                    className="query-editor-add-step__suggestion"
                    title={JSON.stringify(op, null, 2)}
                  >
                    {describeOperation(op)}
                  </button>
                ))
              )}
            </div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
