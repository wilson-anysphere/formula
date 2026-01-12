import { Icon, type IconProps } from "./Icon";

export function ClearIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 7l3-3h7v8H7l-3-3z" />
      <path d="M9 6l2 2" />
      <path d="M11 6l-2 2" />
    </Icon>
  );
}

