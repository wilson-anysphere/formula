import { Icon, type IconProps } from "./Icon";

export function MacroIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M4 2h6l3 3v9H4z" />
      <path d="M10 2v3h3" />
      <path d="M6 9l2-2 2 2 2-2" />
      <path d="M6 11h4" />
    </Icon>
  );
}

