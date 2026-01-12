import { Icon, type IconProps } from "./Icon";

export function ColumnWidthIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 3v8" />
      <path d="M8 3v8" />
      <path d="M12 3v8" />
      <path d="M4 13h8" />
      <polyline points="5.5 12 4 13 5.5 14" />
      <polyline points="10.5 12 12 13 10.5 14" />
    </Icon>
  );
}

