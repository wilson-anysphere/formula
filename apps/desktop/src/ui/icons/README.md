# Desktop UI Icons

This directory contains Formula's internal ribbon/shell icon set.

## Goals

- Clean, consistent look at **16×16** (and scales well to **20×20**).
- No external icon dependencies.
- Icons inherit color via CSS `color` (`currentColor`).

## Usage

```tsx
import { BoldIcon } from "../ui/icons";

<BoldIcon size={16} />
<BoldIcon size={20} style={{ color: "var(--text-primary)" }} />
```

## Adding a new icon

1. Create a new `*.tsx` file in this directory.
2. Render **inline SVG** through the shared `<Icon>` base component:

```tsx
import { Icon, type IconProps } from "./Icon";

export function MyIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="..." />
    </Icon>
  );
}
```

3. Do **not** use hard-coded colors. If you need a filled shape, use:
   - `fill="currentColor"` and optionally `stroke="none"`
4. Export the icon from `index.ts`.

## Ribbon integration

The desktop ribbon renders icons exclusively via the internal ribbon icon library
in `apps/desktop/src/ribbon/icons` using stable `iconId` values in the ribbon
schema.

See `apps/desktop/src/ribbon/icons/README.md` for details.
