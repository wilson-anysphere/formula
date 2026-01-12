import { Icon, type IconProps } from "./Icon";

export function ItalicIcon(props: Omit<IconProps, "children">) {
  return (
    <Icon {...props}>
      <path d="M6 3h6" />
      <path d="M4 13h6" />
      <path d="M10 3l-4 10" />
    </Icon>
  );
}
