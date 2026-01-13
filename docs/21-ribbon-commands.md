# Ribbon commands & controls (developer guide)

This document describes how **ribbon controls** (buttons, toggles, dropdown menu items) are wired to the app’s **command system** and **keybinding system**, and what to update when you add or rename ribbon controls.

Key files:

- Ribbon schema entrypoint (types + tab assembly): `apps/desktop/src/ribbon/ribbonSchema.ts`
- Ribbon tab definitions: `apps/desktop/src/ribbon/schema/*.ts`
- Ribbon mount + wiring: `apps/desktop/src/main.ts` (`mountRibbon(...)`)
- Commands: `apps/desktop/src/extensions/commandRegistry.ts`
- Built-in command registration:
  - `apps/desktop/src/commands/registerBuiltinCommands.ts` (edit/clipboard/view/audit/etc.)
  - `apps/desktop/src/main.ts` (formatting commands like `format.toggleBold`)
- Keybindings: `apps/desktop/src/commands/builtinKeybindings.ts`
- Keybinding dispatcher: `apps/desktop/src/extensions/keybindingService.ts`
- Ribbon UI state overrides: `apps/desktop/src/ribbon/ribbonUiState.ts`
- Keybinding barriers: `apps/desktop/src/keybindingBarrier.js`

---

## Architecture: schema → actions → commands

### 1) Ribbon schema

`apps/desktop/src/ribbon/ribbonSchema.ts` defines the ribbon’s structure (and imports
per-tab definitions from `apps/desktop/src/ribbon/schema/*.ts`):

- Tabs → Groups → Buttons
- Dropdown buttons can contain `menuItems`

Each button/menu item definition contains:

- `id`: **used for wiring** (passed through to `RibbonActions`)
- `testId` (optional): stable E2E hook (becomes `data-testid` on the DOM)

The rendered DOM also includes `data-command-id={button.id}` (see `apps/desktop/src/ribbon/RibbonButton.tsx`) for debugging and unit tests.

### 2) RibbonActions contract

The ribbon itself is “dumb UI”:

- `apps/desktop/src/ribbon/Ribbon.tsx` calls:
  - `actions.onCommand?.(button.id)` for command-like activations (including dropdown menu items)
  - `actions.onToggle?.(button.id, pressed)` for toggle buttons

The host (desktop app) provides these handlers in `apps/desktop/src/main.ts` via `mountRibbon(...)`.

### 3) CommandRegistry is the canonical behavior layer

`apps/desktop/src/extensions/commandRegistry.ts` (`CommandRegistry`) is the canonical place where a command id maps to executable behavior.

When a feature is reachable via both keyboard and ribbon, **both should execute the same CommandRegistry command id** to avoid drift.

Examples of real command ids already in use:

- `clipboard.copy`, `clipboard.paste`, `clipboard.pasteSpecial.values` (registered in `apps/desktop/src/commands/registerBuiltinCommands.ts`)
- `edit.undo`, `edit.redo` (registered in `apps/desktop/src/commands/registerBuiltinCommands.ts`)
- `format.toggleBold`, `format.numberFormat.currency` (registered in `apps/desktop/src/main.ts`)

---

## Guidelines for adding ribbon controls

### Guideline: `RibbonButtonDefinition.id` should be the canonical CommandRegistry id (when functional)

If a ribbon control is wired to do something real, its schema `id` should be the **CommandRegistry command id**.

Why:

- Keyboard shortcuts, command palette, and ribbon can share the same implementation (`CommandRegistry.executeCommand(...)`).
- The command id stays stable even if you move the button between tabs/groups.
- Removes the need for translation layers like `"home.clipboard.copy" → "clipboard.copy"`.

Note: parts of the current schema still use location-based ids (e.g. `home.*`) for historical reasons. As command ids are normalized, some of these ids will change—see `testId` guidance below.

### Guideline: use `testId` for stable E2E hooks

Use `testId` as the stable selector for Playwright / E2E tests:

- `RibbonButton` renders `data-testid={button.testId}` (`apps/desktop/src/ribbon/RibbonButton.tsx`)
- `apps/desktop/test/ribbonTestIds.test.js` enforces that ribbon + backstage `testId`s are **unique** (Playwright strict-mode) and that certain required `testId`s exist.

Do **not** rely on `id`/`data-command-id` for E2E. `id`s may change as commands are normalized.

### Guideline: RibbonActions should route to `CommandRegistry.executeCommand`

In `apps/desktop/src/main.ts`, the ribbon is mounted with:

```ts
mountRibbon(ribbonReactRoot, {
  onCommand: (commandId) => { /* ... */ },
  onToggle: (commandId, pressed) => { /* ... */ },
});
```

Prefer:

- `commandRegistry.executeCommand(commandId)` (or the `executeCommand(...)` helper in `apps/desktop/src/main.ts`)

So keyboard shortcuts (`KeybindingService`) and ribbon share behavior.

If a control is a toggle and you need the **next** pressed state, consider one of these patterns:

1. Make the command accept an argument: `commandRegistry.executeCommand("format.setBold", pressed)`
2. Keep a `format.toggleX` command, and let the command compute the next state from the document/selection.

### Guideline: update ribbon selection format state + UI overrides when ids change

The ribbon UI merges:

- internal UI state (e.g. local toggle state), and
- app-driven overrides from `RibbonUiState` (`apps/desktop/src/ribbon/ribbonUiState.ts`)

`apps/desktop/src/main.ts` maintains these overrides in `scheduleRibbonSelectionFormatStateUpdate()`:

- `pressedById`: drives toggle buttons (Bold/Italic/Underline/Wrap/etc.)
- `labelById`: drives dynamic button labels (e.g. `home.number.numberFormat`)
- `disabledById`: disables controls based on editing/runtime state

If you **rename a ribbon id**, you must update every place that uses that id as a key:

- `scheduleRibbonSelectionFormatStateUpdate()` in `apps/desktop/src/main.ts`
- Any other readers of `RibbonUiState` (e.g. `apps/desktop/src/ribbon/FileBackstage.tsx` reads `uiState.pressedById["file.save.autoSave"]`)
- Unit tests that assert by `data-command-id` (e.g. `apps/desktop/src/ribbon/__tests__/RibbonUiStateOverrides.vitest.ts`)

Tip: after changing an id, run a repo search for the old string.

---

## Keybinding barriers in ribbon menus/backstage (`data-keybinding-barrier`)

Ribbon dropdown menus and the File backstage behave like menus/dialogs: they need arrow keys, Tab, Escape, etc. to be handled locally without global shortcuts stealing those key events.

We use a DOM attribute:

- `data-keybinding-barrier="true"`

Where it’s used today:

- Ribbon dropdown menus: `apps/desktop/src/ribbon/RibbonButton.tsx`
- Ribbon tab overflow menu: `apps/desktop/src/ribbon/Ribbon.tsx`
- File backstage overlay: `apps/desktop/src/ribbon/FileBackstage.tsx`

How it works:

- `KeybindingService` checks for `data-keybinding-barrier="true"` in the event target ancestry and skips dispatching when inside a barrier (see `isInsideKeybindingBarrier` in `apps/desktop/src/extensions/keybindingService.ts`).
- Some non-KeybindingService global listeners use the generic helpers in `apps/desktop/src/keybindingBarrier.js`:
  - `isEventWithinKeybindingBarrier(event)`
  - `markKeybindingBarrier(el)`

When adding a new ribbon popover/menu/backstage-like surface, ensure its root is marked as a keybinding barrier.

---

## Worked example: add a new Format command + ribbon button + keybinding

Goal: add a new formatting command and expose it consistently via:

- CommandRegistry (canonical behavior)
- a keybinding
- a ribbon button with a stable E2E `testId`

Example: **Strikethrough** as a proper command (`format.toggleStrikethrough`) instead of a ribbon-only id.

### 1) Register the command

Add a built-in command registration in `apps/desktop/src/main.ts` near the existing formatting registrations (search for `format.toggleBold`).

Sketch:

```ts
commandRegistry.registerBuiltinCommand(
  "format.toggleStrikethrough",
  t("command.format.toggleStrikethrough"),
  () => {
    const sheetId = app.getCurrentSheetId();
    const selection = app.getSelectionRanges();
    const next = !computeSelectionFormatState(app.getDocument(), sheetId, selection).strikethrough;
    applyFormattingToSelection(
      "Strikethrough",
      (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { font: { strike: next } }, { label: "Strikethrough" });
          if (ok === false) applied = false;
        }
        return applied;
      },
      { forceBatch: true },
    );
  },
  { category: t("commandCategory.format") },
);
```

### 2) Add a keybinding

In `apps/desktop/src/commands/builtinKeybindings.ts`, add a binding that targets the same command id:

```ts
{
  command: "format.toggleStrikethrough",
  key: "ctrl+5",
  mac: "cmd+shift+x",
  when: WHEN_SPREADSHEET_READY,
}
```

### 3) Add the ribbon button (schema)

In `apps/desktop/src/ribbon/ribbonSchema.ts`, add a button (e.g. Home → Font group) whose `id` matches the command id:

```ts
{
  id: "format.toggleStrikethrough",
  label: "Strike",
  ariaLabel: "Strikethrough",
  iconId: "strikethrough",
  kind: "toggle",
  size: "icon",
  testId: "ribbon-strikethrough",
}
```

E2E should select this control via `data-testid="ribbon-strikethrough"`.

### 4) Wire RibbonActions to execute the command

In `apps/desktop/src/main.ts`, in the `mountRibbon(...).onCommand` handler, ensure that activations call:

- `commandRegistry.executeCommand(commandId)` (or the existing `executeCommand(commandId)` helper)

If the `onCommand` handler still uses a large switch statement, add a case for the new command id and route it through `CommandRegistry` rather than duplicating formatting logic.

### 5) Keep toggle state in sync (RibbonUiState)

If the ribbon toggle should reflect selection state (Excel-style), update `scheduleRibbonSelectionFormatStateUpdate()` in `apps/desktop/src/main.ts`:

- Add/update `pressedById["format.toggleStrikethrough"] = formatState.strikethrough`
- If you changed ids, also update any tests that assert `data-command-id` (for example `apps/desktop/src/ribbon/__tests__/RibbonUiStateOverrides.vitest.ts`).

---

## Related tests (helpful when changing ids)

- `apps/desktop/test/ribbonTestIds.test.js` (unique + required `testId`s)
- `apps/desktop/src/ribbon/__tests__/RibbonSchema.vitest.ts` (schema invariants)
- `apps/desktop/src/ribbon/__tests__/RibbonUiStateOverrides.vitest.ts` (pressed/label overrides by id)
