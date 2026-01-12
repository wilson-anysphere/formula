import { Icon, type IconProps } from "./Icon";

export function SyncScrollIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 4v8" />
      <polyline points="6.5 5.5 8 4 9.5 5.5" />
      <polyline points="6.5 10.5 8 12 9.5 10.5" />
      <path d="M4 8h8" />
    </Icon>
  );
}

