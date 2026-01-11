# Formula Design System

Pixel-perfect mockups and design specifications for Formula's UI.

## Viewing Mockups

Open the HTML files directly in a browser:

```bash
open mockups/spreadsheet-main.html
```

## Design Philosophy

### Anti-"AI Slop" Principles

1. **No purple gradients** - We use electric blue (#3b82f6) as primary accent
2. **No Inter/Roboto** - IBM Plex family (Sans + Mono) for distinctive typography
3. **Deep, not flat** - Rich charcoal backgrounds with subtle depth, not gray-on-white
4. **Data-dense** - Optimized for power users who work with large datasets
5. **Sharp, not soft** - Precise borders, defined edges, professional feel

### Aesthetic Direction

**Inspiration:**
- Bloomberg Terminal (data density, professional gravitas)
- Linear (polish, attention to detail)
- Raycast (keyboard-first, focused UI)
- VS Code (developer-grade tooling)

**NOT like:**
- Notion (too playful, too much whitespace)
- Generic SaaS dashboards (purple gradients, rounded everything)
- Google Sheets (utilitarian, no personality)

---

## Color Palette

### Base Colors (Dark Theme)

```css
--bg-base: #0f1114;      /* App background */
--bg-surface: #161a1e;   /* Panels, sidebars */
--bg-elevated: #1c2127;  /* Cards, inputs */
--bg-hover: #252b33;     /* Hover states */
--bg-active: #2d353f;    /* Active/pressed */
```

### Grid Colors

```css
--grid-bg: #12151a;                      /* Cell background */
--grid-line: rgba(255, 255, 255, 0.06);  /* Grid lines */
--grid-header-bg: #181c22;               /* Row/column headers */
--grid-selection: rgba(59, 130, 246, 0.15);
--grid-selection-border: #3b82f6;
```

### Text Hierarchy

```css
--text-primary: #f1f3f5;    /* Main text */
--text-secondary: #8b939e;  /* Labels, descriptions */
--text-muted: #5c6370;      /* Placeholders, hints */
--text-disabled: #3d4450;   /* Disabled states */
```

### Accent & Semantic

```css
--accent: #3b82f6;          /* Primary actions, focus rings */
--accent-hover: #2563eb;    /* Hover state */
--accent-subtle: rgba(59, 130, 246, 0.12);

--success: #10b981;         /* Positive values, confirmations */
--warning: #f59e0b;         /* Warnings */
--error: #ef4444;           /* Errors, negative values */
--info: #06b6d4;            /* Informational */
```

### Formula Syntax

```css
--syntax-function: #60a5fa;   /* SUM, VLOOKUP, etc. */
--syntax-reference: #a78bfa;  /* A1, $B$2, Sheet1!A1 */
--syntax-string: #fbbf24;     /* "text" */
--syntax-number: #34d399;     /* 123, 45.67 */
--syntax-operator: #f1f3f5;   /* +, -, *, / */
--syntax-error: #f87171;      /* #REF!, #VALUE! */
```

---

## Typography

### Font Stack

```css
--font-sans: 'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif;
--font-mono: 'IBM Plex Mono', 'SF Mono', Consolas, monospace;
```

### Sizes

| Use Case | Size | Weight | Font |
|----------|------|--------|------|
| Grid cells | 12px | 400 | Mono |
| UI labels | 13px | 400-500 | Sans |
| Panel titles | 13px | 600 | Sans |
| Section headers | 11px | 600 | Sans |
| Column headers | 11px | 500 | Mono |

### Grid Typography

- Numbers: **right-aligned**, tabular numerals (`font-variant-numeric: tabular-nums`)
- Text: left-aligned
- Headers: centered, uppercase, letter-spacing: 0.02em

---

## Spacing Scale

```css
--space-1: 4px;
--space-2: 8px;
--space-3: 12px;
--space-4: 16px;
--space-5: 20px;
--space-6: 24px;
--space-8: 32px;
```

**Usage:**
- `space-1`: Tight gaps (icon + label)
- `space-2`: Related elements (buttons in a group)
- `space-3`: Standard component padding
- `space-4`: Section margins, larger padding
- `space-6+`: Major section breaks

---

## Component Patterns

### Inputs

```css
/* Base input */
height: 28px;
padding: 0 8px;
background: var(--bg-elevated);
border: 1px solid var(--border-default);
border-radius: 4px;

/* Focus state */
border-color: var(--accent);
box-shadow: 0 0 0 2px var(--accent-subtle);
```

### Buttons

```css
/* Default button */
height: 28px;
padding: 0 12px;
background: var(--bg-elevated);
border: 1px solid var(--border-default);
border-radius: 4px;

/* Primary button */
background: var(--accent);
border-color: var(--accent);
color: white;
```

### Panels

```css
background: var(--bg-surface);
border-left: 1px solid var(--border-subtle); /* or border-right */
```

### Grid Cells

```css
height: 24px;
padding: 0 8px;
border-right: 1px solid var(--grid-line);
border-bottom: 1px solid var(--grid-line);
```

---

## Animation

### Transitions

```css
--transition-fast: 100ms ease;    /* Hover states */
--transition-normal: 150ms ease;  /* Focus, panel opens */
```

### Principles

1. **Subtle, not flashy** - Animations should feel snappy, not theatrical
2. **Purpose-driven** - Only animate to provide feedback or guide attention
3. **Respect reduced-motion** - Check `prefers-reduced-motion`

---

## Icons

- **Size**: 16px default, 14px for tight spaces
- **Stroke**: 2px, round caps and joins
- **Style**: Outlined (not filled)
- **Source**: Lucide icons (MIT licensed)

```html
<svg class="icon" viewBox="0 0 24 24">
  <!-- icon paths -->
</svg>
```

---

## Mockup Files

| File | Description |
|------|-------------|
| `spreadsheet-main.html` | Main app with grid, formula bar, and AI panel |
| *(more coming)* | |

---

## Implementation Notes for Agents

When implementing UI components:

1. **Reference the mockups** - Don't guess at spacing, colors, or typography
2. **Use CSS variables** - All colors and spacing should use the design tokens
3. **Test at 1x and 2x** - Ensure crisp rendering on Retina displays
4. **Check dark mode only** - We're dark-first (light theme is future work)
5. **Match the font stack** - Don't substitute fonts without approval

### CSS Import

```css
/* At top of component CSS */
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
```

### Color Blind Accessibility

- Don't rely solely on red/green for positive/negative
- Use icons or patterns as secondary indicators
- Test with Sim Daltonism or similar tools
