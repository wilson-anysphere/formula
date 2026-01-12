import { Icon, type IconProps } from "./Icon";

export function RowHeightIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M3 4h8" />
      <path d="M3 8h8" />
      <path d="M3 12h8" />
      <path d="M13 4v8" />
      <polyline points="12 5.5 13 4 14 5.5" />
      <polyline points="12 10.5 13 12 14 10.5" />
    </Icon>
  );
}

