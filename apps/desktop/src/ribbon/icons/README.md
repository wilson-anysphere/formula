# Desktop Ribbon Icons

The desktop ribbon renders icons through a small internal icon library.

- Icons are referenced by stable string ids: `RibbonIconId`
- The ribbon schema assigns icons via `iconId` (not by command id)
- Rendering is done through `<RibbonIcon id="..." />`

## Usage

```tsx
import { RibbonIcon } from "./RibbonIcon";

<RibbonIcon id="bold" />;
```

## Adding a new ribbon icon

1. Create an icon component.

   Prefer reusing an existing icon from `apps/desktop/src/ui/icons` when possible.
   For ribbon-specific icons, add a new component to `commonIcons.tsx` and render
   inline SVG via the shared `<Icon>` base component.

2. Register the icon in `RibbonIcon.tsx` by adding it to the `ribbonIcons` map.
   The map key becomes the public `RibbonIconId`, so keep it **stable** and
   **semantic**.

3. Use the new id in `apps/desktop/src/ribbon/ribbonSchema.ts`:

```ts
{
  id: "home.font.bold",
  label: "Bold",
  size: "icon",
  iconId: "bold",
}
```

## Guidelines

- Do not hardcode colors in SVG (use `currentColor`).
- Keep the icon geometry aligned to a `viewBox="0 0 16 16"` for consistency.
