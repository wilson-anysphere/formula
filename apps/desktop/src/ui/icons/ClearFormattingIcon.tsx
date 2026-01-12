import { Icon, type IconProps } from "./Icon";

export function ClearFormattingIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M10.5 3.5l2 2-6 6H4.5l-2-2 6-6z" />
      <path d="M4 13h9" />
    </Icon>
  );
}

