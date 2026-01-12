import { Icon, type IconProps } from "./Icon";

export function LinkIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6.2 6.2L4.8 7.6a2.5 2.5 0 0 0 3.5 3.5l1.4-1.4" />
      <path d="M9.8 9.8l1.4-1.4a2.5 2.5 0 0 0-3.5-3.5L6.3 6.3" />
      <path d="M6.8 9.2l2.4-2.4" />
    </Icon>
  );
}

