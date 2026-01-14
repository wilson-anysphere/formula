# UX Design Principles

## Overview

The user experience must honor Excel's familiar mental model while introducing modern conveniences. Power users should feel faster, not constrained. Novices should feel guided, not overwhelmed.

---

## Design Philosophy

### Core Principles

1. **Familiarity First**: The grid is sacred. Don't reinvent what works.
2. **Keyboard-Driven**: Every action should be accessible without a mouse.
3. **Progressive Disclosure**: Simple by default, powerful when needed.
4. **Immediate Feedback**: Every action should have visible results.
5. **Reversible Actions**: Undo everything, always.

### Anti-Patterns to Avoid

- Modal dialogs that block workflow
- Nested menus more than 2 levels deep
- Settings that require restart
- Features hidden behind right-click only
- Animations that delay user actions

---

## Grid Interface

### Cell Layout

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  Formula Bar                                                                │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ fx │ =SUM(A1:A10) + VLOOKUP(B1, Data!A:C, 3, FALSE)                │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
├─────────────────────────────────────────────────────────────────────────────┤
│     │  A      │  B      │  C      │  D      │  E      │  F      │         │
├─────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│  1  │ Product │ Q1      │ Q2      │ Q3      │ Q4      │ Total   │         │
├─────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│  2  │ Alpha   │ 1,234   │ 2,345   │ 3,456   │ 4,567   │ 11,602  │         │
├─────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│  3  │ Beta    │ 987     │ 1,098   │ 1,209   │ 1,320   │ 4,614   │         │
├─────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│  4  │ Gamma   │ 567     │ 678     │ 789     │ 890     │ 2,924   │         │
├─────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│  5  │ TOTAL   │ 2,788   │ 4,121   │ 5,454   │ 6,777   │ 19,140  │         │
└─────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┘
```

### Visual Hierarchy

| Element | Treatment | Purpose |
|---------|-----------|---------|
| Headers | Bold, subtle background | Identify structure |
| Data | Regular weight | Primary content |
| Formulas | Show calculated value | Users care about results |
| Selection | Blue border, light fill | Current focus |
| Errors | Red background | Draw attention |
| Changes | Yellow flash | Confirm action |

---

## Command Palette

### Trigger: `Cmd+Shift+P` (Mac) / `Ctrl+Shift+P` (Windows/Linux)

**Note:** `Cmd/Ctrl+K` is reserved for **inline AI edit** directly in the grid selection (see `apps/desktop/src/app/spreadsheetApp.ts`), so the command palette uses `Cmd/Ctrl+Shift+P` to avoid a keybinding conflict.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  ┌───────────────────────────────────────────────────────────────────┐     │
│  │ > Insert pivot table                                                │     │
│  └───────────────────────────────────────────────────────────────────┘     │
│                                                                             │
│  RECENT                                                                     │
│  ├── Insert Chart                                              ⌘⇧C         │
│  ├── Format as Currency                                        ⌘⇧$         │
│  └── Sort Descending                                           ⌘⇧↓         │
│                                                                             │
│  SUGGESTIONS                                                                │
│  ├── Insert Pivot Table from A1:F100                                       │
│  ├── Insert Column Before                                                   │
│  └── Insert Row Above                                          ⌘⇧I         │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Command Categories

1. **Navigation**: Go to cell, sheet, named range
2. **Editing**: Insert, delete, format, find/replace
3. **Data**: Sort, filter, pivot, chart
4. **View**: Zoom, freeze, split, hide
5. **AI**: Analyze, explain, generate, transform

Implementation note: PivotTable ownership boundaries (model schema vs compute vs XLSX import/export)
are captured in [ADR-0005](./adr/ADR-0005-pivot-tables-ownership-and-data-flow.md).

### Search Behavior

- Fuzzy matching: "pvt tbl" → "Pivot Table"
- Recent commands prioritized
- Shortcut search: "/" → shows all shortcuts
- Context-aware suggestions *(planned)*

---

## Keyboard Shortcuts

> Note: This document is a UX/design overview. For the authoritative shortcut list in this repo, see
> [`instructions/ui.md`](../instructions/ui.md).

### App

| Shortcut | Action |
|----------|--------|
| `Cmd+Shift+P` / `Ctrl+Shift+P` | Open command palette |
| `/` (in command palette) | Shortcut search |
| `Cmd+K` / `Ctrl+K` | Inline AI edit (transform selection) |

### Navigation
 
| Shortcut | Action |
|----------|--------|
| `Arrow keys` | Move selection |
| `Ctrl+Arrow` | Jump to edge of data |
| `Ctrl+Home` | Go to A1 |
| `Ctrl+End` | Go to last used cell |
| `Cmd/Ctrl+G` | Go to... dialog |
| `Tab` | Move right, wrap to next row |
| `Enter` | Move down, wrap to next column |
| `Page Up/Down` | Scroll viewport |
| `F6` / `Shift+F6` | Cycle focus between ribbon, formula bar, grid, sheet tabs, and status bar |

### Selection

| Shortcut | Action |
|----------|--------|
| `Shift+Arrow` | Extend selection |
| `Ctrl+Shift+Arrow` | Extend to edge of data |
| `Ctrl+A` | Select all cells |
| `Ctrl+Space` | Select entire column |
| `Shift+Space` | Select entire row |
| `Ctrl+Shift+*` (aka `Ctrl+Shift+8` on some keyboards; `Ctrl+*` on the numpad) | Select current region |

### Editing

| Shortcut | Action |
|----------|--------|
| `F2` | Edit cell |
| `Enter` | Confirm and move down |
| `Tab` | Confirm and move right |
| `Escape` | Cancel edit |
| `Delete` | Clear cell contents |
| `Ctrl+;` | Insert current date |
| `Ctrl+Shift+;` | Insert current time |
| `Ctrl+D` | Fill down |
| `Ctrl+R` | Fill right |
| `Cmd/Ctrl+F` | Find |
| `Cmd+Option+F` (Mac) / `Ctrl+H` (Windows/Linux) | Replace |
| `Cmd/Ctrl+G` | Go to… |

### Formatting

| Shortcut | Action |
|----------|--------|
| `Cmd/Ctrl+B` | Bold |
| `Ctrl+I` | Italic |
| `Cmd/Ctrl+U` | Underline |
| `Cmd/Ctrl+1` | Format cells dialog |
| `Cmd/Ctrl+Shift+$` | Currency format |
| `Cmd/Ctrl+Shift+%` | Percentage format |
| `Cmd/Ctrl+Shift+#` | Date format |

### AI (New)

| Shortcut | Action |
|----------|--------|
| `Cmd+K` (Mac) / `Ctrl+K` (Windows/Linux) | Inline AI edit |
| `Cmd+I` (Mac) / `Ctrl+Shift+A` (Windows/Linux) | Toggle AI chat sidebar |
| `Tab` (in formula bar, when an AI suggestion is shown) | Accept AI suggestion |

Platform notes:

- **macOS:** `Cmd+I` is reserved for **AI Chat**. Use `Ctrl+I` for **Italic** (Excel-compatible).
- **Windows/Linux:** `Ctrl+I` is reserved for **Italic** (Excel-compatible). Use `Ctrl+Shift+A` to toggle the AI chat sidebar.

---

## Formula Bar

### Enhanced Formula Editing

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  A1        ▼  │ fx │ =IF(                                                  │
│               │    │   SUM(B1:B10) > 1000,                                 │
│               │    │   "Over Budget",                                       │
│               │    │   "Within Budget"                                      │
│               │    │ )                                                      │
├─────────────────────────────────────────────────────────────────────────────┤
│  PARAMETERS                                                                 │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │ IF(logical_test, [value_if_true], [value_if_false])                │   │
│  │     ↳ SUM(B1:B10) > 1000  →  TRUE                                  │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Features

1. **Expandable**: Grows vertically for complex formulas
2. **Syntax Highlighting**: Functions, references, operators colored
3. **Auto-Indentation**: Nested functions indented
4. **Parameter Hints**: Shows function signature and current argument
5. **Range Preview**: Hover shows range contents
6. **AI Suggestions**: Tab-completion inline

---

## Formula Debugging

### Step-Through Debugger

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  FORMULA DEBUGGER                                               [×] Close   │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  =VLOOKUP(A1, Data!A:C, 3, FALSE)                                          │
│                                                                             │
│  STEP 1: Evaluate A1                                                        │
│  ├── A1 = "Product-123"                                                    │
│  │                                                                          │
│  STEP 2: Evaluate Data!A:C                                                 │
│  ├── Range: Data!A1:C50 (50 rows)                                          │
│  │   ┌─────────────┬──────────┬─────────┐                                  │
│  │   │ Product-123 │ Widgets  │ $19.99  │ ← Match found                    │
│  │   │ Product-456 │ Gadgets  │ $29.99  │                                  │
│  │   │ ...         │ ...      │ ...     │                                  │
│  │   └─────────────┴──────────┴─────────┘                                  │
│  │                                                                          │
│  STEP 3: Look up column 3                                                   │
│  ├── Column 3 value = $19.99                                               │
│  │                                                                          │
│  RESULT: $19.99                                                             │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Error Explanation

When a formula returns an error:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  ⚠️ #N/A Error in D5                                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  Formula: =VLOOKUP(A5, B1:C10, 2, FALSE)                                   │
│                                                                             │
│  PROBLEM: The lookup value "XYZ-999" was not found in the first column    │
│  of the range B1:C10.                                                       │
│                                                                             │
│  SUGGESTIONS:                                                               │
│  • Check if "XYZ-999" exists in column B                                   │
│  • Verify the lookup range is correct                                       │
│  • Consider using IFERROR to handle missing values                         │
│                                                                             │
│  [Fix with AI]  [Show lookup range]  [Ignore]                              │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Context Menus

### Cell Context Menu

```
┌──────────────────────────────┐
│ Cut                    ⌘X    │
│ Copy                   ⌘C    │
│ Paste                  ⌘V    │
│ Paste Special...       ⌘⇧V   │
├──────────────────────────────┤
│ Insert...                    │
│ Delete...                    │
│ Clear Contents         Del   │
├──────────────────────────────┤
│ Format Cells...        ⌘1    │
│ Column Width...              │
│ Row Height...                │
├──────────────────────────────┤
│ 🤖 Ask AI about this...     │
│ 🤖 Fill similar cells...    │
├──────────────────────────────┤
│ Add Comment            ⇧F2   │
│ View History...              │
└──────────────────────────────┘
```

### Selection-Aware Options

The context menu adapts to what's selected:

- **Single cell**: Standard options
- **Range with data**: Sort, filter, chart options
- **Table header**: Column operations
- **Formula cell**: Debug, explain options
- **Error cell**: Fix suggestions

---

## Panels

### AI Chat Panel

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  AI ASSISTANT                                              [−] [□] [×]      │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  [You] What's the trend in column C?                                        │
│                                                                             │
│  [AI] Looking at the data in C2:C100, I can see:                           │
│                                                                             │
│  📈 **Upward trend** with 15% growth over the period                       │
│                                                                             │
│  • Starting value (C2): $1,234                                              │
│  • Ending value (C100): $1,419                                              │
│  • Peak: $1,567 in row 78                                                   │
│  • Average: $1,298                                                          │
│                                                                             │
│  [Create trend chart]  [Show details]  [Add forecast]                      │
│                                                                             │
├─────────────────────────────────────────────────────────────────────────────┤
│  ┌───────────────────────────────────────────────────────────────────┐     │
│  │ Ask a question about your data...                           Send │     │
│  └───────────────────────────────────────────────────────────────────┘     │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Version History Panel

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  VERSION HISTORY                                           [−] [□] [×]      │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  TODAY                                                                      │
│  ├── 3:45 PM - You                                                         │
│  │   Updated Q4 forecasts (Sheet1: D10:D50)                                │
│  │   [Restore] [Compare]                                                   │
│  │                                                                          │
│  ├── 2:30 PM - Alice                                                       │
│  │   Added new product rows                                                │
│  │   [Restore] [Compare]                                                   │
│  │                                                                          │
│  YESTERDAY                                                                  │
│  ├── 5:15 PM - Bob                                                         │
│  │   Fixed formula error in totals                                         │
│  │   [Restore] [Compare]                                                   │
│  │                                                                          │
│  CHECKPOINTS                                                                │
│  ├── ★ Q3 Budget Approved - Oct 1                                          │
│  │   Created by: Finance Team                                              │
│  │   [Restore] [Compare]                                                   │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Sheet Tabs (Workbook Navigation)

The workbook uses an Excel-style **sheet tab strip** at the bottom of the window (above the status bar). Tabs are the primary way to navigate and manage worksheets.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  ◀ ▶  Sheet1   Sheet2   Sheet3   [+]                              ⋯         │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Core Interactions (Excel-like)

1. **Create sheet**: Click `+` to create a new sheet.
   - Default naming: `Sheet1`, `Sheet2`, … using the next available number.
   - Insert position: directly after the currently active sheet.

2. **Rename sheet**: Double-click a tab (or `Rename` from context menu).
   - Inline editing on the tab.
   - Validation (match Excel constraints):
     - Unique (case-insensitive) within the workbook.
     - Max length: 31 characters.
     - Disallow: `: \\ / ? * [ ]`.

3. **Reorder sheets**: Drag tabs to reorder.
   - Auto-scroll tab strip while dragging near edges.
   - Reorder does not affect formulas (formulas reference sheets by name), but the new order must persist to storage and XLSX.

4. **Delete sheet**: Context menu → `Delete`.
   - Prevent deleting the last remaining sheet.
   - If a deleted sheet is referenced by formulas, Excel turns those references into `#REF!` (behavior to emulate).

5. **Hide / Unhide sheets**:
   - Hide: Context menu → `Hide`.
     - Prevent hiding the last *visible* sheet.
   - Unhide: Context menu on the tab strip background → `Unhide…` (shows a list).
     - Only `hidden` sheets appear.
     - `veryHidden` sheets are preserved on XLSX round-trip but not exposed in the standard UI (Excel requires VBA).

6. **Tab colors**:
   - Show sheet color on the tab (underline or fill).
   - Preserve colors from XLSX on load/save.
   - Optional: context menu → `Tab Color…` for a picker/palette.

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+PgUp` / `Cmd+PgUp` | Activate previous visible sheet (wrap around) |
| `Ctrl+PgDn` / `Cmd+PgDn` | Activate next visible sheet (wrap around) |

---

## Status Bar

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  Ready │ Sum: 19,140 │ Avg: 3,828 │ Count: 5 │ 100% │ Sheet 1 of 3 │ ↕     │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Elements

| Element | Description |
|---------|-------------|
| Status | Current mode (Ready, Edit, etc.) |
| Quick stats | Sum, Average, Count of selection |
| Zoom | Click to adjust |
| Sheet navigation | Current sheet position |
| View controls | Scroll lock, page breaks |

---

## Notifications

### Toast Notifications

```
┌─────────────────────────────────────────────────────┐
│  ✓ Saved to cloud                            [×]   │
│    Last saved: just now                            │
└─────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────┐
│  ⚠️ Alice is editing cells you're viewing    [×]   │
│    [See their changes]                             │
└─────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────┐
│  📊 Calculation complete                     [×]   │
│    98,432 cells recalculated in 0.3s              │
└─────────────────────────────────────────────────────┘
```

### Inline Notifications

For cell-level issues:

```
     │  A      │  B      │
├────┼─────────┼─────────┤
│  1 │ 100     │ #DIV/0! │ ← Hover shows: "Division by zero in formula =A1/C1"
│  2 │ 200     │ 20      │
```

---

## Responsive Design

### Window Size Adaptations

| Width | Adaptation |
|-------|------------|
| < 800px | Hide ribbon, use hamburger menu |
| 800-1200px | Collapsed ribbon groups |
| > 1200px | Full ribbon with labels |

### Panel Behavior

- Panels can be docked left, right, or bottom
- Panels can be floating
- Panels remember position per-document
- Double-click header to maximize

---

## Accessibility

### Screen Reader Support

- All cells have ARIA labels
- Regions announced on navigation
- Selection changes announced
- Error messages read automatically

### Keyboard-Only Usage

- Tab order follows logical flow
- Focus indicators always visible
- No mouse-only interactions
- Shortcuts work in all contexts

### Visual Accessibility

- High contrast mode support
- Minimum 4.5:1 contrast ratios
- No color-only indicators
- Scalable UI (up to 200%)

---

## Loading States

### Initial Load

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                                                                             │
│                        ┌─────────────────────┐                              │
│                        │    📊 Formula       │                              │
│                        │                     │                              │
│                        │ Loading workbook... │                              │
│                        │ ████████░░░░ 67%    │                              │
│                        │                     │                              │
│                        │ Parsing formulas... │                              │
│                        └─────────────────────┘                              │
│                                                                             │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Calculation Progress

For long calculations, show progress in status bar:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  Calculating... │ ████████░░ 80% │ 78,432 / 98,432 cells │ ETA: 2s        │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Empty States

### New Workbook

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                        Welcome to Formula                                   │
│                                                                             │
│              Start typing in any cell, or try these:                       │
│                                                                             │
│              📁 Open a file                    ⌘O                          │
│              📊 Import from Excel              ⌘⇧I                         │
│              🤖 Ask AI to create something                                 │
│              📋 Paste data from clipboard      ⌘V                          │
│                                                                             │
│              Recent files:                                                  │
│              • Budget 2024.xlsx                                            │
│              • Sales Report Q3.xlsx                                        │
│              • Inventory.csv                                                │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Theming

Formula is designed **light-first**. The default theme preference is **Light**, with optional **Dark**, **System**, and **High Contrast** modes.

### Theme switching

- **Ribbon:** `View → Theme` (System / Light / Dark / High Contrast)
- **Command palette:** search for `Theme` and run either:
  - `Theme…` (opens a picker)
  - `Theme: System` / `Theme: Light` / `Theme: Dark` / `Theme: High Contrast`

Theme preference is persisted (desktop: `localStorage` key `formula.settings.appearance.v1`) and applied on startup by the theme controller (see `apps/desktop/src/theme/themeController.js`) via the `data-theme` attribute on `<html>`.

### Light Theme (Default)

```css
:root {
  --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  --font-mono: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
  --bg-primary: #ffffff;
  --bg-secondary: #f5f5f5;
  --bg-tertiary: #e8e8e8;
  --text-primary: #1a1a1a;
  --text-secondary: #666666;
  --border: #d4d4d4;
  --accent: #0066cc;
  --accent-light: #e6f0ff;
  --link: #0969da;
  --error: #d32f2f;
  --warning: #ed6c02;
  --success: #2e7d32;
}
```

### Dark Theme

```css
:root[data-theme="dark"] {
  --bg-primary: #1e1e1e;
  --bg-secondary: #252526;
  --bg-tertiary: #333333;
  --text-primary: #e4e4e4;
  --text-secondary: #a0a0a0;
  --border: #404040;
  --accent: #4da6ff;
  --accent-light: #1a3a5c;
  --link: #4da6ff;
  --error: #f44336;
  --warning: #ff9800;
  --success: #4caf50;
}
```

### System preference resolution (`Theme: System`)

When the user selects **System**, the app resolves the active theme using media queries:

- `forced-colors` / higher-contrast preferences → `data-theme="high-contrast"`
- otherwise `prefers-color-scheme: dark` → `data-theme="dark"`
- otherwise → `data-theme="light"`

This is centralized in `ThemeController` so changes propagate live without requiring a restart.
If you need to drive theme changes programmatically, use the controller API (the ribbon wires to this):

```typescript
const themeController = new ThemeController();
themeController.start(); // applies persisted preference (default: light)

// Opt-in to following OS changes.
themeController.setThemePreference("system");
```
