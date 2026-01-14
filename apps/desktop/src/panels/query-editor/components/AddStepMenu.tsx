import React, { useEffect, useMemo, useRef, useState } from "react";

import { stableStringify, type ArrowTableAdapter, type DataTable, type Query, type QueryOperation } from "@formula/power-query";
import { getLocale, t } from "../../../i18n/index.js";
import { formatQueryOperationLabel } from "../operationLabels";

function prettyJson(value: unknown): string {
  try {
    return JSON.stringify(JSON.parse(stableStringify(value)), null, 2);
  } catch (err) {
    try {
      return JSON.stringify(value, null, 2);
    } catch {
      return String(value);
    }
  }
}

export function AddStepMenu(props: {
  onAddStep: (op: QueryOperation) => void;
  onAiSuggest?: (intent: string, ctx: { query: Query; preview: DataTable | ArrowTableAdapter | null }) => Promise<QueryOperation[]>;
  aiContext: { query: Query; preview: DataTable | ArrowTableAdapter | null };
}) {
  const [intent, setIntent] = useState("");
  // Keep a synchronous copy of the latest intent so we can respond to key events
  // even when React state updates from `input` events haven't flushed yet.
  const intentRef = useRef("");
  const [suggestions, setSuggestions] = useState<QueryOperation[] | null>(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);
  const aiRequestIdRef = useRef(0);

  const [menuOpen, setMenuOpen] = useState(false);
  const menuRootRef = useRef<HTMLDivElement | null>(null);
  const menuTriggerRef = useRef<HTMLButtonElement | null>(null);

  const locale = getLocale();

  const schema = props.aiContext.preview?.columns ?? [];
  const columnNames = useMemo(
    () =>
      schema
        .map((col) => String(col?.name ?? "").trim())
        .filter((name) => name.length > 0),
    [schema],
  );
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
  const schemaRequiredHint = t("queryEditor.addStep.schemaRequiredHint");

  const menuGroups = useMemo(() => {
    const schemaRequiredDisabled = !schemaReady;
    const schemaItem = (item: Omit<MenuItem, "disabled" | "disabledReason">): MenuItem => ({
      ...item,
      disabled: schemaRequiredDisabled,
      disabledReason: schemaRequiredDisabled ? schemaRequiredReason : undefined,
    });

    const uniqueName = (base: string, used: Set<string>): string => {
      let name = base;
      let suffix = 1;
      while (used.has(name)) {
        name = `${base} ${suffix}`;
        suffix += 1;
      }
      return name;
    };
    const usedColumnNames = new Set(columnNames);
    for (const step of props.aiContext.query.steps) {
      const op = step.operation;
      if (op.type === "addColumn" && typeof op.name === "string") {
        const trimmed = op.name.trim();
        if (trimmed) usedColumnNames.add(trimmed);
      }
      if (op.type === "renameColumn" && typeof op.newName === "string") {
        const trimmed = op.newName.trim();
        if (trimmed) usedColumnNames.add(trimmed);
      }
      if (op.type === "unpivot") {
        const nameColumn = typeof op.nameColumn === "string" ? op.nameColumn.trim() : "";
        const valueColumn = typeof op.valueColumn === "string" ? op.valueColumn.trim() : "";
        if (nameColumn) usedColumnNames.add(nameColumn);
        if (valueColumn) usedColumnNames.add(valueColumn);
      }
      if (op.type === "groupBy" && Array.isArray(op.aggregations)) {
        for (const agg of op.aggregations) {
          const name =
            typeof agg?.as === "string" && agg.as.trim()
              ? agg.as.trim()
              : typeof agg?.op === "string" && typeof agg?.column === "string"
                ? (() => {
                    const opName = agg.op.trim();
                    const column = agg.column.trim();
                    if (!opName || !column) return null;
                    return `${opName} of ${column}`;
                  })()
                : null;
          if (name) usedColumnNames.add(name);
        }
      }
    }

    return [
      {
        id: "rows",
        label: t("queryEditor.addStep.category.rows"),
        items: [
          {
            id: "take",
            label: t("queryEditor.addStep.op.keepTopRows"),
            create: () => ({ type: "take", count: 100 }),
          },
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
          {
            id: "distinctRows",
            label: t("queryEditor.addStep.op.removeDuplicates"),
            create: () => ({ type: "distinctRows", columns: null }),
          },
          {
            id: "removeRowsWithErrors",
            label: t("queryEditor.addStep.op.removeRowsWithErrors"),
            create: () => ({ type: "removeRowsWithErrors", columns: null }),
          },
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
            create: () => {
              const oldName = firstColumnName;
              const baseNewName = `${oldName || "Column"} (Renamed)`;
              const used = new Set(usedColumnNames);
              used.delete(oldName);
              const newName = uniqueName(baseNewName, used);
              return { type: "renameColumn", oldName, newName };
            },
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
            create: () => {
              const groupColumns = [firstColumnName];
              const baseAggName = t("queryEditor.addStep.op.groupBy.count");
              const as = uniqueName(baseAggName, new Set(groupColumns));
              return {
                type: "groupBy",
                groupColumns,
                aggregations: [{ column: secondColumnName, op: "count", as }],
              };
            },
          }),
          schemaItem({
            id: "unpivot",
            label: t("queryEditor.addStep.op.unpivot"),
            create: () => {
              // Prefer unpivoting a "value" column (2nd col) while leaving the first
              // column as an identifier, mirroring typical Power Query usage.
              const unpivotColumn = columnNames.length > 1 ? secondColumnName : firstColumnName;
              const nameColumn = uniqueName("Attribute", usedColumnNames);
              const used = new Set(usedColumnNames);
              used.add(nameColumn);
              const valueColumn = uniqueName("Value", used);
              return { type: "unpivot", columns: [unpivotColumn], nameColumn, valueColumn };
            },
          }),
          schemaItem({
            id: "fillDown",
            label: t("queryEditor.addStep.op.fillDown"),
            create: () => ({ type: "fillDown", columns: [firstColumnName] }),
          }),
          schemaItem({
            id: "replaceValues",
            label: t("queryEditor.addStep.op.replaceValues"),
            create: () => ({ type: "replaceValues", column: firstColumnName, find: "", replace: "" }),
          }),
          schemaItem({
            id: "addColumn",
            label: t("queryEditor.addStep.op.addColumn"),
            create: () => ({
              type: "addColumn",
              name: uniqueName("Custom", usedColumnNames),
              formula: firstColumnName ? `[${firstColumnName}]` : "0",
            }),
          }),
        ] satisfies MenuItem[],
      },
    ];
  }, [locale, columnNames, props.aiContext.query.steps, firstColumnName, schemaReady, schemaRequiredReason, secondColumnName]);

  useEffect(() => {
    if (!menuOpen) return;
    if (typeof document === "undefined") return;

    const focusTrigger = () => {
      // Only refocus the trigger for keyboard-driven close actions (e.g. Escape).
      // Avoid stealing focus on outside clicks, where the user likely clicked an
      // input/button elsewhere.
      queueMicrotask(() => menuTriggerRef.current?.focus());
    };

    const onMouseDown = (evt: MouseEvent) => {
      const target = evt.target as Node | null;
      if (!target) return;
      const root = menuRootRef.current;
      if (!root) return;
      if (root.contains(target)) return;
      setMenuOpen(false);
    };

    const onKeyDown = (evt: KeyboardEvent) => {
      if (evt.key === "Escape") {
        setMenuOpen(false);
        focusTrigger();
      }
    };

    document.addEventListener("mousedown", onMouseDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onMouseDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [menuOpen]);

  const handleMenuKeyDown: React.KeyboardEventHandler<HTMLDivElement> = (evt) => {
    if (
      evt.key !== "ArrowDown" &&
      evt.key !== "ArrowUp" &&
      evt.key !== "Home" &&
      evt.key !== "End" &&
      evt.key !== "Enter"
    )
      return;
    const root = menuRootRef.current;
    if (!root) return;
    const items = Array.from(
      root.querySelectorAll<HTMLButtonElement>(".query-editor-add-step__menu-popover button:not(:disabled)"),
    );
    if (items.length === 0) return;

    const active = document.activeElement;
    const currentIdx = items.findIndex((el) => el === active);

    if (evt.key === "Enter") {
      if (active && active instanceof HTMLButtonElement && items.includes(active)) {
        evt.preventDefault();
        active.click();
      }
      return;
    }

    let nextIdx = 0;
    if (evt.key === "Home") {
      nextIdx = 0;
    } else if (evt.key === "End") {
      nextIdx = items.length - 1;
    } else if (evt.key === "ArrowDown") {
      nextIdx = currentIdx >= 0 ? (currentIdx + 1) % items.length : 0;
    } else if (evt.key === "ArrowUp") {
      nextIdx = currentIdx >= 0 ? (currentIdx - 1 + items.length) % items.length : items.length - 1;
    }

    evt.preventDefault();
    items[nextIdx]?.focus();
  };

  useEffect(() => {
    if (!menuOpen) return;
    const root = menuRootRef.current;
    if (!root) return;
    // Ensure focus moves into the menu for keyboard users (best-effort).
    // Avoid focusing back to the trigger on close so we don't steal focus from
    // outside-click targets (e.g. the AI intent input).
    queueMicrotask(() => {
      const first = root.querySelector<HTMLButtonElement>(".query-editor-add-step__menu-popover button:not(:disabled)");
      first?.focus();
    });
  }, [menuOpen]);

  const updateIntent = (next: string) => {
    // Avoid double-handling when both `input` and React's normalized `change` fire.
    if (next === intentRef.current) return;
    intentRef.current = next;
    setIntent(next);
    if (aiLoading) {
      // Treat edits during loading as a cancellation of the in-flight request:
      // ignore its eventual response and stop showing a loading state.
      aiRequestIdRef.current += 1;
      setAiLoading(false);
    }
    if (aiError) setAiError(null);
    if (suggestions) setSuggestions(null);
  };

  async function runAiSuggest(overrideIntent?: string): Promise<void> {
    const trimmed = (overrideIntent ?? intentRef.current ?? intent).trim();
    if (!trimmed) return;
    if (!props.onAiSuggest) return;
    const requestId = (aiRequestIdRef.current += 1);
    setAiLoading(true);
    setAiError(null);
    setSuggestions(null);
    try {
      const ops = await props.onAiSuggest(trimmed, props.aiContext);
      if (requestId !== aiRequestIdRef.current) return;
      setSuggestions(ops ?? []);
    } catch (err) {
      if (requestId !== aiRequestIdRef.current) return;
      setAiError(err instanceof Error ? err.message : String(err));
      setSuggestions(null);
    } finally {
      // Only update loading state if this is still the latest request.
      if (requestId === aiRequestIdRef.current) setAiLoading(false);
    }
  }

  return (
    <div className="query-editor-add-step">
      <div ref={menuRootRef} className="query-editor-add-step__menu">
        <button
          type="button"
          className="query-editor-add-step__menu-trigger"
          onClick={() => setMenuOpen((open) => !open)}
          aria-haspopup="menu"
          aria-expanded={menuOpen}
          ref={menuTriggerRef}
        >
          {t("queryEditor.addStep.addStep")}
        </button>
        {menuOpen ? (
          <div className="query-editor-add-step__menu-popover" role="menu" onKeyDown={handleMenuKeyDown}>
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
                        // Menu selection should return focus to the trigger (menu-button pattern).
                        queueMicrotask(() => menuTriggerRef.current?.focus());
                      }}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>
              </div>
            ))}
            {!schemaReady ? <div className="query-editor-add-step__menu-hint">{schemaRequiredHint}</div> : null}
          </div>
        ) : null}
      </div>

      {props.onAiSuggest ? (
        <div>
          <input
            value={intent}
            // React normalizes `onChange` for text inputs to fire on `input` events, but
            // some tests dispatch only `input` and rely on immediate availability of the
            // typed value. Handle both to be robust.
            onChange={(e) => updateIntent(e.currentTarget.value)}
            onInput={(e) => updateIntent((e.currentTarget as HTMLInputElement).value)}
            onKeyDown={(e) => {
              if (e.key !== "Enter") return;
              if (aiLoading) return;
              const trimmed = e.currentTarget.value.trim();
              if (!trimmed) return;
              e.preventDefault();
              void runAiSuggest(trimmed);
            }}
            placeholder={t("queryEditor.addStep.aiPlaceholder")}
            className="query-editor-add-step__ai-input"
          />
          <button
            type="button"
            onClick={() => {
              void runAiSuggest();
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
                    onClick={() => {
                      props.onAddStep(op);
                      setSuggestions(null);
                    }}
                    className="query-editor-add-step__suggestion"
                    title={prettyJson(op)}
                  >
                    {formatQueryOperationLabel(op)}
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
