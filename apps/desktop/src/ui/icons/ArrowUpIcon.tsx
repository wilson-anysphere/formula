import { Icon, type IconProps } from "./Icon";

export function ArrowUpIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M8 12V4" />
      <polyline points="6.5 5.5 8 4 9.5 5.5" />
    </Icon>
  );
}

