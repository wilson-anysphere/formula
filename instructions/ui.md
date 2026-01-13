# Workstream B: UI/UX (TypeScript/React)

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Build the presentation layer: canvas-based grid renderer, formula bar, command palette, panels, and theming. Achieve **60fps scrolling with millions of rows** while maintaining full visual fidelity.

**The goal:** Excel functionality with Cursor polish. **Light mode by default**. Professional. Power-user focused.

---

## Scope

### Your Code

| Location | Purpose |
|----------|---------|
| `packages/grid` | Canvas-based grid renderer, virtualization |
| `packages/spreadsheet-frontend` | High-level spreadsheet UI coordination |
| `packages/text-layout` | Text measurement, font handling |
| `apps/desktop/src/grid/` | Grid integration, cell editing |
| `apps/desktop/src/formula-bar/` | Formula bar, autocomplete |
| `apps/desktop/src/panels/` | Side panels, dialogs |
| `apps/desktop/src/layout/` | Window layout, panes |
| `apps/desktop/src/theme/` | Theming system |
| `apps/desktop/src/styles/` | Global CSS |

### Your Documentation

- **Primary:** [`docs/03-rendering-ui.md`](../docs/03-rendering-ui.md) — canvas rendering, virtualization, overlays
- **UX Design:** [`docs/12-ux-design.md`](../docs/12-ux-design.md) — interface design, keyboard shortcuts
- **Ribbon wiring:** [`docs/21-ribbon-commands.md`](../docs/21-ribbon-commands.md) — ribbon schema ids, CommandRegistry, keybindings, and keybinding barriers

### Design Reference

```bash
mockups/spreadsheet-main.html    # Main app layout
mockups/ai-agent-mode.html       # Agent execution view
mockups/command-palette.html     # Quick command search
mockups/README.md                # Full design system
```

> ⚠️ **Mockups are directional, not literal.**
> Use them for vision and design language, but apply judgment—add polish, fix inconsistencies, implement missing interactions. Follow the **principles** over exact pixels.

---

## Key Requirements

### Canvas Grid Rendering

1. **Full grid on Canvas** — no DOM cells (Google Sheets approach)
2. **Virtualized scrolling** — O(v) complexity (v = visible cells only)
3. **Batch draw calls** — minimize GPU context switches
4. **Device pixel ratio awareness** — crisp on Retina displays
5. **Smooth scrolling** — handle 33M pixel browser limit (~1M rows at 30px)
6. **Layered canvases:** Grid → Content → Selection → DOM Overlays

### Coordinate Systems

```
Screen Space → Canvas Space → Cell Space → Data Space
```

- Handle frozen rows/columns
- Support split panes
- Efficient hit testing

### Formula Bar

- Always visible
- AI autocomplete integration
- Syntax highlighting
- Error display with location

### Excel Functionality (Non-Negotiable)

- **Full ribbon interface** with all Excel buttons
- **Sheet tabs** at bottom
- **Status bar** with Sum/Avg/Count
- **Hide / Unhide rows and columns** (shared-grid parity)
- **All Excel keyboard shortcuts** (F2, F4, Ctrl+D, etc.)
- **Context menus** matching Excel behavior

--- 

## Design Rules

### ⛔ CRITICAL: FUNCTIONALITY ≠ STYLING ⛔

**These are TWO INDEPENDENT AXES:**

```
FUNCTIONALITY (# of Excel buttons/features)
        ↑
  LOW   │   HIGH
────────┼────────→ STYLING (clean vs bloated)
        │
```

**THE RULE:**
1. **ADD all Excel buttons** — Font, Bold, Italic, Borders, Fill, Merge, AutoSum, ALL of them
2. **STYLE them cleanly** — No Microsoft bloat, just clean Cursor-style buttons

**CORRECT:** HIGH functionality + CLEAN styling

**FAILURE MODES:**
- ❌ Microsoft Bureaucracy: HIGH functionality + BLOATED styling
- ❌ Stripped Minimal: LOW functionality + CLEAN styling

### Design Tokens

```css
/* Backgrounds */
--bg-app: #f8f8f8;
--bg-surface: #ffffff;
--bg-hover: #f0f0f0;

/* Text */
--text-primary: #1f1f1f;
--text-secondary: #6e6e6e;

/* Borders */
--border: #e5e5e5;
--border-strong: #d0d0d0;

/* Accent */
--accent: #0969da;

/* Grid */
--row-height: 22px;
--col-width: 90px;
--header-width: 40px;
```

### Typography

- **Cells:** Monospace (SFMono-Regular, Menlo, Consolas)
- **UI:** System fonts (-apple-system, "Segoe UI")
- **Sizes:** 10px, 11px, 12px, 13px, 14px

---

## Keyboard Shortcuts

### Excel-Compatible (MUST work)

| Shortcut | Action |
|----------|--------|
| `F2` | Edit cell |
| `Shift+F2` | Add comment (open comments panel and focus new comment input) |
| `F4` | Toggle absolute/relative reference |
| `Tab` | Move selection right (wrap within selection) / commit edit and move right |
| `Shift+Tab` | Move selection left (wrap within selection) / commit edit and move left |
| `Enter` | Move selection down (wrap within selection) / commit edit and move down |
| `Shift+Enter` | Move selection up (wrap within selection) / commit edit and move up |
| `Ctrl+D` | Fill down |
| `Ctrl+;` | Insert date |
| `Alt+=` | AutoSum |
| `Ctrl+C/V/X/Z/Y` | Standard clipboard/undo |
| `Cmd/Ctrl+F` | Find |
| `Cmd+Option+F` (macOS) / `Ctrl+H` (Windows/Linux) | Replace |
| `Cmd/Ctrl+G` | Go to |
| `Ctrl+PgUp` (Windows/Linux) / `Cmd+PgUp` (macOS) | Activate previous visible sheet (wrap around) |
| `Ctrl+PgDn` (Windows/Linux) / `Cmd+PgDn` (macOS) | Activate next visible sheet (wrap around) |

### App Shortcuts

| Shortcut | Action |
|----------|--------|
| `Cmd/Ctrl+Shift+P` | Open command palette |
| `Cmd/Ctrl+Shift+M` | Toggle comments panel |
| `F6` / `Shift+F6` | Cycle keyboard focus between ribbon, formula bar, grid, sheet tabs, and status bar |
| `/` (in command palette) | Search shortcuts/keybindings |

### AI Shortcuts

| Shortcut | Action |
|----------|--------|
| `Cmd/Ctrl+K` | Inline AI edit |
| `Cmd+I` (macOS) / `Ctrl+Shift+A` (Windows/Linux) | Toggle AI chat sidebar |
| `Tab` (in formula bar, when an AI suggestion is shown) | Accept AI suggestion |

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Scroll FPS | 60fps with 1M+ rows |
| Cold start | <1 second to interactive grid |
| Keystroke latency | <16ms |

---

## Build & Run

```bash
# Install dependencies
pnpm install

# Run desktop app (dev server)
pnpm dev:desktop
# Opens at http://localhost:4174

# Run grid package in isolation
pnpm --dir packages/grid dev

# Type check
pnpm typecheck
```

---

## Coordination Points

- **Core Engine Team:** You consume their WASM API for calculations
- **File I/O Team:** You display what they parse
- **AI Team:** AI sidebar, inline edit UI, suggestion display
- **Collaboration Team:** Presence indicators, cursor display

---

## Testing

```bash
# Unit tests
pnpm test

# E2E tests (Playwright)
pnpm test:e2e

# Visual tests with Xvfb (headless)
xvfb-run --auto-servernum pnpm test:e2e
```

---

## Accessibility

- Full ARIA labels for screen readers
- Complete keyboard navigation
- High contrast mode (respect system preference)
- Font scaling (respect system font size)
- Reduced motion (disable animations when requested)
