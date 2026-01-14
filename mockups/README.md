# Formula Design System

> **Excel functionality + Cursor polish**
> 
> Full Excel feature set (ribbon, formulas, sheets) with modern, clean aesthetics.
> Light mode. Professional. Polished.

---

## ⛔ STOP. READ THIS FIRST. NON-NEGOTIABLE. ⛔

**FUNCTIONALITY and STYLING are TWO INDEPENDENT AXES.**

```
FUNCTIONALITY (# of Excel features/buttons)
        ↑
        │
  LOW   │   HIGH ← we want HIGH functionality
────────┼────────→ STYLING
        │         (clean vs bloated)
        │
        we want CLEAN styling →
```

### THE RULE:

**1. ADD ALL THE EXCEL BUTTONS.** Every feature Excel has. Font dropdown. Bold. Italic. Underline. Borders. Fill. Merge. AutoSum. Conditional Formatting. ALL OF THEM.

**2. STYLE THEM CLEANLY.** No Microsoft bloat. No complex nested layouts. No heavy borders. No gradient icons. Just clean, minimal, Cursor-style buttons.

### THE FAILURE MODES (both wrong):

| Mode | Functionality | Styling | Why it's wrong |
|------|--------------|---------|----------------|
| **Microsoft Bureaucracy** | HIGH ✓ | BLOATED ✗ | Complex layouts, heavy UI, enterprise feel |
| **Stripped Minimal** | LOW ✗ | CLEAN ✓ | Removed features to "look minimal" |

### THE CORRECT MODE:

| Mode | Functionality | Styling | Result |
|------|--------------|---------|--------|
| **Formula** | HIGH ✓ | CLEAN ✓ | Full Excel power, Cursor aesthetics |

### "MINIMAL" MEANS:

- ✅ Visual simplicity (clean borders, subtle colors, tight spacing)
- ❌ NOT feature removal (removing buttons to look "clean")

### EXAMPLES:

```
Want a font dropdown?     → ADD IT. Style it clean.
Want Bold/Italic buttons? → ADD THEM. Style them clean.  
Want Conditional Format?  → ADD IT. Style it clean.
Want Merge & Center?      → ADD IT. Style it clean.
Want all 8 ribbon tabs?   → ADD THEM. Style them clean.
```

**DO NOT swing between "Microsoft bloat" and "stripped down". Stay in the HIGH functionality + CLEAN styling quadrant. ALWAYS.**

---

## Philosophy

### What We Are
- **An Excel replacement** - all the features power users expect
- **With modern polish** - clean, light, professional like Cursor
- **AI-native** - deeply integrated, not bolted on

### What We're Not
- A "modern minimal" reimagining that loses functionality
- A dark-mode-only developer tool
- A chatbot with a spreadsheet attached

### Two Modes, Two Densities

**Spreadsheet View** (where you work)
- Dense, tight, every pixel counts
- Ribbon is compact but readable
- Grid is information-dense
- AI sidebar is minimal

**Agent View** (where you observe)
- Relaxed, spacious, easy to scan
- Steps are clearly separated
- Code blocks have breathing room
- This is a view, not a workspace

## Target Users

CFOs, financial analysts, accountants, operations managers, data scientists. People who:
- Use Excel 8+ hours/day
- Know every keyboard shortcut
- Build complex financial models
- Need their tools to be reliable and professional

## Design Tokens

### Core Principle: Density + Depth

Every pixel counts. Tighten spacing, add subtle shades for depth. No wasted whitespace.

### Colors

```css
/* Backgrounds - layered with subtle depth */
--bg-app: #f5f5f5;          /* App background */
--bg-surface: #fafafa;      /* Panels, card backgrounds */
--bg-elevated: #ffffff;     /* Active content, inputs */
--bg-hover: #eeeeee;        /* Hover states */
--bg-inset: #f0f0f0;        /* Inset panels, code blocks */

/* Text - clear hierarchy */
--text-primary: #1a1a1a;    /* Primary text */
--text-secondary: #5c5c5c;  /* Secondary text */
--text-tertiary: #8a8a8a;   /* Hints, labels */

/* Borders */
--border: #e0e0e0;          /* Standard border */
--border-strong: #c8c8c8;   /* Emphasized border */

/* Accent - professional blue */
--accent: #0969da;          /* Primary accent */
--link: #0969da;            /* Hyperlinks */
--accent-bg: #ddf4ff;       /* Accent background */
--accent-border: #54aeff;   /* Accent border */

/* Selection */
--selection-bg: rgba(9, 105, 218, 0.08);
--selection-border: #0969da;

/* Semantic */
--green: #1a7f37;           /* Positive values */
--green-bg: #dafbe1;
--red: #cf222e;             /* Negative values */
--red-bg: #ffebe9;
--yellow: #9a6700;          /* Warnings */
```

### Typography

```css
/* System fonts - professional, readable */
--font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
--font-mono: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;

/* Sizes - compact, dense */
9px   /* Smallest labels, hints */
10px  /* Labels, tags, secondary */
11px  /* Cells, UI text */
12px  /* Primary UI, inputs */
```

### Sizing

```css
/* Border radius - tight, modern */
--radius: 4px;      /* Panels, cards, buttons */
--radius-sm: 3px;   /* Small elements, tags */

/* Heights - compact */
20px  /* Small buttons, grid cells */
22px  /* Standard small buttons */
24px  /* Primary buttons */
26px  /* Large inputs */

/* Grid */
Row height: 20px
Column width: 80px (min)
Row header: 36px

/* Ribbon */
Tab height: 26px
Content height: 52px
```

## Layout

### App Structure

```
┌─────────────────────────────────────────────────────────────┐
│ Title Bar (window controls, app name, doc name, actions)    │
├─────────────────────────────────────────────────────────────┤
│ Ribbon (Home | Insert | Formulas | Data | ...)              │
│ ┌─────────┬─────────┬─────────┬─────────┬─────────┐        │
│ │ Clipboard│  Font   │ Number  │  Cells  │ Editing │        │
│ └─────────┴─────────┴─────────┴─────────┴─────────┘        │
├─────────────────────────────────────────────────────────────┤
│ Formula Bar (name box | fx | formula input | AI trigger)    │
├───────────────────────────────────────────────┬─────────────┤
│                                               │             │
│                                               │ AI Sidebar  │
│                    Grid                       │ - Modes     │
│                                               │ - Context   │
│                                               │ - Messages  │
│                                               │ - Input     │
├───────────────────────────────────────────────┴─────────────┤
│ Sheet Tabs (nav | Sheet1 | Sheet2 | ... | +)                │
├─────────────────────────────────────────────────────────────┤
│ Status Bar (status | sum | average | count | zoom)          │
└─────────────────────────────────────────────────────────────┘
```

### Ribbon Groups

**Home**: Clipboard, Font, Alignment, Number, Cells, Editing
**Insert**: Tables, Charts, Sparklines, Filters, Links, Text
**Formulas**: Function Library, Named Ranges, Formula Auditing
**Data**: Get Data, Sort & Filter, Data Tools, Outline
**Review**: Comments, Changes, Protection
**View**: Workbook Views, Show, Zoom, Window

## AI Integration

> **This is a Cursor product.** All AI goes through Cursor servers.
> No local models. No API keys. No provider selection.
> Cursor controls the harness and prompts.

### One Panel, Full Power

The AI sidebar is a unified panel - no mode tabs. Just type what you want:
- Ask questions → AI answers
- Request changes → AI proposes diff  
- Complex tasks → opens Agent view

### Key Features

- **Context awareness** - shows what sheets/ranges AI sees
- **Diff preview** - shows proposed changes before apply
- **Accept/reject** - per-change and bulk controls
- **Agent mode** - full-screen view for autonomous multi-step execution

### Triggers

- **⌘K** - Inline edit (from anywhere)
- **Tab** - Accept AI suggestion (in formula bar)
- **⌘I** (macOS) / **Ctrl+Shift+A** (Windows/Linux) - Toggle AI sidebar
- **Agent view** - Separate full-height view for autonomous tasks

## Keyboard Shortcuts

### Excel-Compatible (must work)

| Shortcut | Action |
|----------|--------|
| F2 | Edit cell |
| F4 | Toggle absolute/relative |
| Ctrl+C/X/V | Copy/Cut/Paste |
| Ctrl+Z/Y | Undo/Redo |
| Ctrl+D | Fill down |
| Ctrl+R | Fill right |
| Ctrl+; | Insert date |
| Alt+= | AutoSum |
| Ctrl+Home | Go to A1 |
| Ctrl+End | Go to last cell |
| Ctrl+Arrow | Jump to edge |

### AI Shortcuts

| Shortcut | Action |
|----------|--------|
| ⌘K | Open AI inline edit |
| ⌘I (macOS) / Ctrl+Shift+A (Windows/Linux) | Toggle AI sidebar |
| Tab | Accept AI suggestion |

## Mockup Files

| File | Description |
|------|-------------|
| `spreadsheet-main.html` | Main app - ribbon, grid, AI sidebar |
| `ai-agent-mode.html` | Agent execution view |
| `command-palette.html` | Quick command/function search |

> ⚠️ **These are directional mockups, not pixel-perfect specs.**
>
> Use them for **vision, layout, and design language** — but apply judgment:
> - They are **crude prototypes** with missing features and rough edges
> - Many interactions, states, and edge cases are not shown
> - Polish, refine, and fill gaps as you implement
> - Follow the **principles** (Excel functionality + Cursor polish) over exact pixels

## Implementation Notes

### Do

- Use the ribbon - it's what Excel users expect
- Keep light mode - professionals prefer it
- Show all status bar calculations
- Use monospace for cells/formulas
- Maintain Excel keyboard shortcuts
- Show formula bar always
- **Tighten all spacing** - no wasted whitespace
- **Layer backgrounds** - use subtle shades for depth
- **Keep elements small** - 20-24px heights, 9-12px fonts

### Don't

- Hide functionality in menus
- Use dark mode by default
- Simplify away features
- Change keyboard shortcuts
- Make AI the primary interface
- **Add giant buttons** - no "Chat" "Edit" "Agent" pills
- **Use gradients on icons** - solid accent colors only
- **Pad excessively** - every pixel has value
- **Make headings oversized** - content over chrome

### Anti-patterns (AI Slop)

These make designs look generic and unpolished:

```
❌ Giant mode buttons with icons
❌ Gradient backgrounds on small elements
❌ 16-20px padding everywhere
❌ 14-16px font sizes for UI text
❌ Large rounded corners (8px+)
❌ "Chat" / "Edit" / "Agent" as big pills
❌ Oversized avatars and icons
```

---

*Excel functionality. Cursor density. Every pixel designed.*
